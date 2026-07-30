using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    /// <summary>
    /// Requests naming suggestions only when the user configured an API key.
    /// </summary>
    internal sealed class NlpSuggestionClient
    {
        internal const string Endpoint =
            "https://gpt.uni-muenster.de/v1/chat/completions";

        private static readonly HttpClient Client = CreateClient();

        public async Task<string> SuggestAsync(
            NlpShapeCandidate candidate, string apiKey)
        {
            if (candidate == null)
                throw new ArgumentNullException(nameof(candidate));
            if (string.IsNullOrWhiteSpace(apiKey))
                return string.Empty;

            string prompt = BuildPrompt(candidate);
            var requestBody = new
            {
                model = "Llama-3.3-70B",
                messages = new[]
                {
                    new
                    {
                        role = "system",
                        content = "You review PASS process-model labels. "
                            + "Return three concise alternative labels, "
                            + "one per line, without commentary."
                    },
                    new { role = "user", content = prompt }
                },
                temperature = 0.3,
                max_tokens = 120
            };

            using (HttpRequestMessage request = new HttpRequestMessage(
                HttpMethod.Post, Endpoint))
            {
                request.Headers.Authorization =
                    new AuthenticationHeaderValue("Bearer", apiKey);
                request.Content = new StringContent(
                    JsonConvert.SerializeObject(requestBody),
                    Encoding.UTF8,
                    "application/json");

                using (HttpResponseMessage response =
                    await Client.SendAsync(request))
                {
                    string responseBody =
                        await response.Content.ReadAsStringAsync();
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "Suggestion service returned HTTP "
                            + (int)response.StatusCode + ".");
                    }

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

        private static HttpClient CreateClient()
        {
            return new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(45)
            };
        }

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
