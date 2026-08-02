using ALPS_Visio_AddIn_rewrite.NlpChecking;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace ALPS_Visio_AddIn_rewrite.Tests.NlpChecking
{
    [TestClass]
    public sealed class NlpSuggestionClientTests
    {
        [TestMethod]
        public async Task SuggestAsync_OpenAiProtocol_SendsBearerChatRequest()
        {
            RecordingHandler handler = new RecordingHandler(
                "{\"choices\":[{\"message\":{\"content\":\"Approve request\\nReview request\"}}]}");
            NlpSuggestionClient client = Client(handler);
            NlpShapeCandidate candidate = Candidate();

            string result = await client.SuggestAsync(
                candidate, Provider(NlpApiProtocol.OpenAiCompatible));

            Assert.AreEqual("Approve request\nReview request", result);
            Assert.AreEqual(HttpMethod.Post, handler.Method);
            Assert.AreEqual(
                "https://provider.example/v1/chat/completions",
                handler.RequestUri.ToString());
            Assert.AreEqual("Bearer", handler.AuthorizationScheme);
            Assert.AreEqual("test-key", handler.AuthorizationParameter);
            StringAssert.Contains(handler.Body, "Current label: bad label");
            StringAssert.Contains(handler.Body, "FullySpecifiedSubject");
            Assert.IsFalse(handler.Body.Contains("Confidential page"));
            Assert.IsFalse(handler.Body.Contains("Internal shape name"));
        }

        [TestMethod]
        public async Task SuggestAsync_AnthropicProtocol_SendsVersionedMessagesRequest()
        {
            RecordingHandler handler = new RecordingHandler(
                "{\"content\":[{\"type\":\"text\",\"text\":\"Customer service\"}]}");
            NlpSuggestionClient client = Client(handler);

            string result = await client.SuggestAsync(
                Candidate(), Provider(NlpApiProtocol.AnthropicCompatible));

            Assert.AreEqual("Customer service", result);
            Assert.AreEqual(
                "https://provider.example/v1/messages",
                handler.RequestUri.ToString());
            Assert.AreEqual("test-key", handler.Header("x-api-key"));
            Assert.AreEqual("2023-06-01", handler.Header("anthropic-version"));
            Assert.IsNull(handler.AuthorizationScheme);
            StringAssert.Contains(handler.Body, "\"system\"");
        }

        [TestMethod]
        public async Task GetModelsAsync_OpenAiProtocol_SortsAndDeduplicatesModels()
        {
            RecordingHandler handler = new RecordingHandler(
                "{\"data\":[{\"id\":\"z-model\"},{\"id\":\"A-model\"},{\"id\":\"a-model\"}]}");
            NlpSuggestionClient client = Client(handler);

            IList<string> models = await client.GetModelsAsync(
                Provider(NlpApiProtocol.OpenAiCompatible));

            CollectionAssert.AreEqual(
                new[] { "A-model", "z-model" }, models.ToArray());
            Assert.AreEqual(HttpMethod.Get, handler.Method);
            Assert.AreEqual(
                "https://provider.example/v1/models",
                handler.RequestUri.ToString());
        }

        [TestMethod]
        public async Task GetModelsAsync_AnthropicProtocol_UsesPagedModelsEndpoint()
        {
            RecordingHandler handler = new RecordingHandler(
                "{\"data\":[{\"id\":\"claude-test\"}]}");
            NlpSuggestionClient client = Client(handler);

            await client.GetModelsAsync(
                Provider(NlpApiProtocol.AnthropicCompatible));

            Assert.AreEqual(
                "https://provider.example/v1/models?limit=1000",
                handler.RequestUri.ToString());
            Assert.AreEqual("2023-06-01", handler.Header("anthropic-version"));
        }

        [TestMethod]
        public async Task GetModelsAsync_NonJsonError_DoesNotExposeResponseBody()
        {
            RecordingHandler handler = new RecordingHandler(
                "private upstream token", HttpStatusCode.BadGateway);
            NlpSuggestionClient client = Client(handler);

            InvalidOperationException exception =
                await Assert.ThrowsExactlyAsync<InvalidOperationException>(
                    () => client.GetModelsAsync(
                        Provider(NlpApiProtocol.OpenAiCompatible)));

            StringAssert.Contains(exception.Message, "HTTP 502");
            Assert.IsFalse(exception.Message.Contains("private upstream token"));
        }

        private static NlpSuggestionClient Client(RecordingHandler handler)
        {
            return new NlpSuggestionClient(new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(2)
            });
        }

        private static NlpProviderSettings Provider(NlpApiProtocol protocol)
        {
            return new NlpProviderSettings
            {
                BaseUrl = "https://provider.example/v1/",
                ApiKey = "test-key",
                Model = "test-model",
                Protocol = protocol,
                IsBuiltIn = true
            };
        }

        private static NlpShapeCandidate Candidate()
        {
            return new NlpShapeCandidate
            {
                PageName = "Confidential page",
                ShapeId = 42,
                ShapeName = "Internal shape name",
                Label = "bad label",
                ShapeType = "FullySpecifiedSubject"
            };
        }

        private sealed class RecordingHandler : HttpMessageHandler
        {
            private readonly string responseBody;
            private readonly HttpStatusCode statusCode;
            private readonly Dictionary<string, string> headers =
                new Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase);

            public RecordingHandler(
                string responseBody,
                HttpStatusCode statusCode = HttpStatusCode.OK)
            {
                this.responseBody = responseBody;
                this.statusCode = statusCode;
            }

            public HttpMethod Method { get; private set; }
            public Uri RequestUri { get; private set; }
            public string Body { get; private set; } = string.Empty;
            public string AuthorizationScheme { get; private set; }
            public string AuthorizationParameter { get; private set; }

            public string Header(string name)
            {
                return headers.TryGetValue(name, out string value)
                    ? value
                    : null;
            }

            protected override async Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request,
                CancellationToken cancellationToken)
            {
                Method = request.Method;
                RequestUri = request.RequestUri;
                AuthorizationScheme =
                    request.Headers.Authorization?.Scheme;
                AuthorizationParameter =
                    request.Headers.Authorization?.Parameter;
                foreach (KeyValuePair<string, IEnumerable<string>> header
                    in request.Headers)
                {
                    headers[header.Key] = string.Join(",", header.Value);
                }

                if (request.Content != null)
                    Body = await request.Content.ReadAsStringAsync();

                return new HttpResponseMessage(statusCode)
                {
                    Content = new StringContent(responseBody)
                };
            }
        }
    }
}
