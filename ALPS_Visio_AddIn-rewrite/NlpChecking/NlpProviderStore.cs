using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    /// <summary>
    /// Stores all provider profiles as one DPAPI-protected JSON document.
    /// </summary>
    internal sealed class NlpProviderStore
    {
        internal const string OpenAiId = "openai";
        internal const string AnthropicId = "anthropic";
        internal const string UniGptId = "unigpt";

        private readonly string configurationPath;
        private readonly string legacyApiKeyPath;

        public NlpProviderStore()
        {
            string directory = Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "ALPS Visio Add-In");
            configurationPath = Path.Combine(
                directory, "nlp-providers.dat");
            legacyApiKeyPath = Path.Combine(
                directory, "nlp-api-key.dat");
        }

        public NlpProviderConfiguration Load()
        {
            NlpProviderConfiguration configuration =
                ReadConfiguration() ?? CreateDefaults();
            EnsureBuiltInProviders(configuration);
            MigrateLegacyApiKey(configuration);
            return configuration;
        }

        public void Save(NlpProviderConfiguration configuration)
        {
            if (configuration == null)
                throw new ArgumentNullException(nameof(configuration));

            EnsureBuiltInProviders(configuration);
            Validate(configuration);

            string directory =
                Path.GetDirectoryName(configurationPath);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            string json = JsonConvert.SerializeObject(configuration);
            byte[] clearBytes = Encoding.UTF8.GetBytes(json);
            byte[] protectedBytes = ProtectedData.Protect(
                clearBytes, null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(configurationPath, protectedBytes);
        }

        private NlpProviderConfiguration ReadConfiguration()
        {
            try
            {
                if (!File.Exists(configurationPath))
                    return null;

                byte[] protectedBytes =
                    File.ReadAllBytes(configurationPath);
                byte[] clearBytes = ProtectedData.Unprotect(
                    protectedBytes, null,
                    DataProtectionScope.CurrentUser);
                return JsonConvert.DeserializeObject<
                    NlpProviderConfiguration>(
                        Encoding.UTF8.GetString(clearBytes));
            }
            catch (Exception exception) when (
                exception is IOException
                || exception is UnauthorizedAccessException
                || exception is CryptographicException
                || exception is JsonException)
            {
                System.Diagnostics.Debug.WriteLine(
                    "NLP provider configuration could not be read: "
                    + exception);
                return null;
            }
        }

        private void MigrateLegacyApiKey(
            NlpProviderConfiguration configuration)
        {
            if (!File.Exists(legacyApiKeyPath))
                return;

            try
            {
                byte[] clearBytes = ProtectedData.Unprotect(
                    File.ReadAllBytes(legacyApiKeyPath),
                    null,
                    DataProtectionScope.CurrentUser);
                string legacyKey =
                    Encoding.UTF8.GetString(clearBytes).Trim();
                NlpProviderSettings uniGpt =
                    configuration.Providers.First(provider =>
                        string.Equals(provider.Id, UniGptId,
                            StringComparison.OrdinalIgnoreCase));
                if (string.IsNullOrWhiteSpace(uniGpt.ApiKey))
                    uniGpt.ApiKey = legacyKey;

                Save(configuration);
                File.Delete(legacyApiKeyPath);
            }
            catch (Exception exception) when (
                exception is IOException
                || exception is UnauthorizedAccessException
                || exception is CryptographicException)
            {
                System.Diagnostics.Debug.WriteLine(
                    "Legacy NLP API key could not be migrated: "
                    + exception);
            }
        }

        private static NlpProviderConfiguration CreateDefaults()
        {
            return new NlpProviderConfiguration
            {
                ActiveProviderId = UniGptId,
                Providers = CreateBuiltInProviders()
            };
        }

        private static List<NlpProviderSettings>
            CreateBuiltInProviders()
        {
            return new List<NlpProviderSettings>
            {
                new NlpProviderSettings
                {
                    Id = OpenAiId,
                    DisplayName = "OpenAI",
                    BaseUrl = "https://api.openai.com/v1",
                    Protocol = NlpApiProtocol.OpenAiCompatible,
                    Model = "gpt-4o-mini",
                    IsBuiltIn = true,
                    KnownModels = new List<string>
                    {
                        "gpt-4o-mini"
                    }
                },
                new NlpProviderSettings
                {
                    Id = AnthropicId,
                    DisplayName = "Anthropic",
                    BaseUrl = "https://api.anthropic.com/v1",
                    Protocol =
                        NlpApiProtocol.AnthropicCompatible,
                    Model = "claude-sonnet-4-20250514",
                    IsBuiltIn = true,
                    KnownModels = new List<string>
                    {
                        "claude-sonnet-4-20250514"
                    }
                },
                new NlpProviderSettings
                {
                    Id = UniGptId,
                    DisplayName = "UniGPT",
                    BaseUrl = "https://gpt.uni-muenster.de/v1",
                    Protocol = NlpApiProtocol.OpenAiCompatible,
                    Model = "Llama-3.3-70B",
                    IsBuiltIn = true,
                    KnownModels = new List<string>
                    {
                        "Llama-3.3-70B",
                        "mistral-small",
                        "gpt-oss-120b"
                    }
                }
            };
        }

        private static void EnsureBuiltInProviders(
            NlpProviderConfiguration configuration)
        {
            if (configuration.Providers == null)
            {
                configuration.Providers =
                    new List<NlpProviderSettings>();
            }

            foreach (NlpProviderSettings template
                in CreateBuiltInProviders())
            {
                NlpProviderSettings existing =
                    configuration.Providers.FirstOrDefault(provider =>
                        string.Equals(provider.Id, template.Id,
                            StringComparison.OrdinalIgnoreCase));
                if (existing == null)
                {
                    configuration.Providers.Add(template);
                    continue;
                }

                existing.DisplayName = template.DisplayName;
                existing.BaseUrl = template.BaseUrl;
                existing.Protocol = template.Protocol;
                existing.IsBuiltIn = true;
                if (existing.KnownModels == null
                    || existing.KnownModels.Count == 0)
                {
                    existing.KnownModels = template.KnownModels;
                }
                if (string.IsNullOrWhiteSpace(existing.Model))
                    existing.Model = template.Model;
            }

            if (string.IsNullOrWhiteSpace(
                configuration.ActiveProviderId)
                || configuration.GetActiveProvider() == null)
            {
                configuration.ActiveProviderId = UniGptId;
            }
        }

        private static void Validate(
            NlpProviderConfiguration configuration)
        {
            HashSet<string> ids =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> names =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (NlpProviderSettings provider
                in configuration.Providers)
            {
                if (provider == null
                    || string.IsNullOrWhiteSpace(provider.Id)
                    || !ids.Add(provider.Id))
                {
                    throw new InvalidOperationException(
                        "Provider identifiers must be unique.");
                }

                if (string.IsNullOrWhiteSpace(provider.DisplayName))
                {
                    throw new InvalidOperationException(
                        "Every provider requires a name.");
                }
                provider.DisplayName =
                    provider.DisplayName.Trim();
                if (!names.Add(provider.DisplayName))
                {
                    throw new InvalidOperationException(
                        "Provider names must be unique.");
                }

                if (!Enum.IsDefined(
                    typeof(NlpApiProtocol), provider.Protocol))
                {
                    throw new InvalidOperationException(
                        "The provider API protocol is invalid.");
                }

                if (!Uri.TryCreate(provider.BaseUrl,
                        UriKind.Absolute, out Uri baseUri)
                    || (baseUri.Scheme != Uri.UriSchemeHttps
                        && baseUri.Scheme != Uri.UriSchemeHttp))
                {
                    throw new InvalidOperationException(
                        "Every provider requires a valid HTTP or HTTPS base URL.");
                }

                provider.ApiKey = provider.ApiKey?.Trim();
                provider.Model = provider.Model?.Trim();
                provider.BaseUrl =
                    provider.BaseUrl.Trim().TrimEnd('/');
                provider.KnownModels = (provider.KnownModels
                        ?? new List<string>())
                    .Where(model => !string.IsNullOrWhiteSpace(model))
                    .Select(model => model.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(model => model,
                        StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
        }
    }
}
