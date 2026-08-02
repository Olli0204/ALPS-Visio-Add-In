using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    /// <summary>
    /// Queries provider models and requests suggestions through either the
    /// OpenAI-compatible or Anthropic-compatible API protocol.
    /// </summary>
    internal sealed class NlpSuggestionClient
    {
        private static readonly HttpClient SharedClient = CreateClient();
        private readonly HttpClient client;

        public NlpSuggestionClient()
            : this(SharedClient)
        {
        }

        internal NlpSuggestionClient(HttpClient client)
        {
            this.client = client
                ?? throw new ArgumentNullException(nameof(client));
        }

        public async Task<IList<string>> GetModelsAsync(
            NlpProviderSettings provider)
        {
            ValidateProvider(provider, requireModel: false);
            string relativePath =
                provider.Protocol
                    == NlpApiProtocol.AnthropicCompatible
                    ? "models?limit=1000"
                    : "models";

            using (HttpRequestMessage request = CreateRequest(
                provider, HttpMethod.Get, relativePath))
            using (HttpResponseMessage response =
                await client.SendAsync(request))
            {
                string responseBody =
                    await response.Content.ReadAsStringAsync();
                EnsureSuccess(response, responseBody);

                JObject json = JObject.Parse(responseBody);
                List<string> models = json
                    .SelectTokens("data[*].id")
                    .Select(token => (string)token)
                    .Where(model =>
                        !string.IsNullOrWhiteSpace(model))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(model => model,
                        StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (models.Count == 0)
                {
                    throw new InvalidOperationException(
                        "The provider returned no selectable models.");
                }

                return models;
            }
        }

        public async Task<string> SuggestAsync(
            NlpShapeCandidate candidate,
            NlpProviderSettings provider)
        {
            if (candidate == null)
                throw new ArgumentNullException(nameof(candidate));
            ValidateProvider(provider, requireModel: true);

            string prompt = BuildPrompt(candidate);
            if (provider.Protocol
                == NlpApiProtocol.AnthropicCompatible)
            {
                return await SuggestWithAnthropicAsync(
                    provider, prompt);
            }

            return await SuggestWithOpenAiAsync(provider, prompt);
        }

        private async Task<string> SuggestWithOpenAiAsync(
            NlpProviderSettings provider, string prompt)
        {
            var requestBody = new
            {
                model = provider.Model,
                messages = new[]
                {
                    new
                    {
                        role = "system",
                        content = SystemPrompt
                    },
                    new { role = "user", content = prompt }
                },
                temperature = 0.3,
                max_tokens = 120
            };

            using (HttpRequestMessage request = CreateRequest(
                provider, HttpMethod.Post, "chat/completions"))
            {
                request.Content = new StringContent(
                    JsonConvert.SerializeObject(requestBody),
                    Encoding.UTF8,
                    "application/json");

                using (HttpResponseMessage response =
                    await client.SendAsync(request))
                {
                    string responseBody =
                        await response.Content.ReadAsStringAsync();
                    EnsureSuccess(response, responseBody);

                    JObject json = JObject.Parse(responseBody);
                    string content = (string)json.SelectToken(
                        "choices[0].message.content");
                    if (string.IsNullOrWhiteSpace(content))
                    {
                        throw new InvalidOperationException(
                            "Suggestion service returned no suggestions.");
                    }

                    return content.Trim();
                }
            }
        }

        private async Task<string> SuggestWithAnthropicAsync(
            NlpProviderSettings provider, string prompt)
        {
            var requestBody = new
            {
                model = provider.Model,
                max_tokens = 120,
                temperature = 0.3,
                system = SystemPrompt,
                messages = new[]
                {
                    new { role = "user", content = prompt }
                }
            };

            using (HttpRequestMessage request = CreateRequest(
                provider, HttpMethod.Post, "messages"))
            {
                request.Content = new StringContent(
                    JsonConvert.SerializeObject(requestBody),
                    Encoding.UTF8,
                    "application/json");

                using (HttpResponseMessage response =
                    await client.SendAsync(request))
                {
                    string responseBody =
                        await response.Content.ReadAsStringAsync();
                    EnsureSuccess(response, responseBody);

                    JObject json = JObject.Parse(responseBody);
                    string content = (string)json.SelectToken(
                        "content[0].text");
                    if (string.IsNullOrWhiteSpace(content))
                    {
                        throw new InvalidOperationException(
                            "Suggestion service returned no suggestions.");
                    }

                    return content.Trim();
                }
            }
        }

        private static HttpRequestMessage CreateRequest(
            NlpProviderSettings provider,
            HttpMethod method,
            string relativePath)
        {
            Uri endpoint = new Uri(
                provider.BaseUrl.TrimEnd('/') + "/"
                + relativePath.TrimStart('/'),
                UriKind.Absolute);
            HttpRequestMessage request =
                new HttpRequestMessage(method, endpoint);

            if (provider.Protocol
                == NlpApiProtocol.AnthropicCompatible)
            {
                if (!string.IsNullOrWhiteSpace(provider.ApiKey))
                    request.Headers.Add("x-api-key", provider.ApiKey);
                request.Headers.Add(
                    "anthropic-version", "2023-06-01");
            }
            else if (!string.IsNullOrWhiteSpace(provider.ApiKey))
            {
                request.Headers.Authorization =
                    new AuthenticationHeaderValue(
                        "Bearer", provider.ApiKey);
            }

            return request;
        }

        private static void EnsureSuccess(
            HttpResponseMessage response, string responseBody)
        {
            if (response.IsSuccessStatusCode)
                return;

            string providerMessage = null;
            try
            {
                providerMessage = (string)JObject
                    .Parse(responseBody)
                    .SelectToken("error.message");
            }
            catch (JsonException)
            {
                // Do not expose arbitrary response bodies in the UI.
            }

            throw new InvalidOperationException(
                "Provider returned HTTP "
                + (int)response.StatusCode
                + (string.IsNullOrWhiteSpace(providerMessage)
                    ? "."
                    : ": " + providerMessage));
        }

        private static void ValidateProvider(
            NlpProviderSettings provider, bool requireModel)
        {
            if (provider == null)
                throw new ArgumentNullException(nameof(provider));
            if (string.IsNullOrWhiteSpace(provider.BaseUrl))
            {
                throw new InvalidOperationException(
                    "The provider base URL is missing.");
            }
            if (provider.IsBuiltIn
                && string.IsNullOrWhiteSpace(provider.ApiKey))
            {
                throw new InvalidOperationException(
                    "The provider API key is missing.");
            }
            if (requireModel
                && string.IsNullOrWhiteSpace(provider.Model))
            {
                throw new InvalidOperationException(
                    "No provider model is selected.");
            }
        }

        private static HttpClient CreateClient()
        {
            // VSTO runs inside the .NET Framework CLR, whose process-wide
            // default may still negotiate obsolete TLS versions. Current
            // provider endpoints require TLS 1.2 or newer.
            ServicePointManager.SecurityProtocol |=
                SecurityProtocolType.Tls12;

            HttpClientHandler handler = new HttpClientHandler
            {
                UseCookies = false
            };
            return new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(45)
            };
        }

        private const string SystemPrompt =
            "You review PASS process-model labels. "
            + "Return three concise alternative labels, "
            + "one per line, without commentary.";

        private static string BuildPrompt(NlpShapeCandidate candidate)
        {
            string guidance;
            switch (candidate.ShapeType)
            {
                case "FullySpecifiedSubject":
                case "InterfaceSubject":
                case "MultiSubject":
                    guidance = "Use a role or actor noun phrase.";
                    break;
                case "DoState":
                    guidance = "Use a short verb-object activity phrase.";
                    break;
                case "SendState":
                case "ReceiveState":
                    guidance = "Use a short communication activity phrase.";
                    break;
                case "DoTransition":
                    guidance = "Use a concise condition or event.";
                    break;
                case "MessageSpecification":
                    guidance = "Use a concise business-object noun phrase.";
                    break;
                default:
                    guidance = "Use a concise PASS-appropriate label.";
                    break;
            }

            return "Shape type: " + candidate.ShapeType
                + "\nCurrent label: " + candidate.Label
                + "\nGuidance: " + guidance;
        }
    }
}
