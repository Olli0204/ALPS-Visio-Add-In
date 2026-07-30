using System;
using System.Collections.Generic;
using System.Linq;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    internal enum NlpApiProtocol
    {
        OpenAiCompatible,
        AnthropicCompatible
    }

    internal sealed class NlpProviderSettings
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string BaseUrl { get; set; }
        public NlpApiProtocol Protocol { get; set; }
        public string ApiKey { get; set; }
        public string Model { get; set; }
        public bool IsBuiltIn { get; set; }
        public List<string> KnownModels { get; set; } =
            new List<string>();

        public bool CanRequestSuggestions =>
            !string.IsNullOrWhiteSpace(BaseUrl)
            && !string.IsNullOrWhiteSpace(Model)
            && (!IsBuiltIn || !string.IsNullOrWhiteSpace(ApiKey));

        public NlpProviderSettings Clone()
        {
            return new NlpProviderSettings
            {
                Id = Id,
                DisplayName = DisplayName,
                BaseUrl = BaseUrl,
                Protocol = Protocol,
                ApiKey = ApiKey,
                Model = Model,
                IsBuiltIn = IsBuiltIn,
                KnownModels = KnownModels == null
                    ? new List<string>()
                    : new List<string>(KnownModels)
            };
        }

        public override string ToString()
        {
            return DisplayName ?? string.Empty;
        }
    }

    internal sealed class NlpProviderConfiguration
    {
        public string ActiveProviderId { get; set; }
        public List<NlpProviderSettings> Providers { get; set; } =
            new List<NlpProviderSettings>();

        public NlpProviderSettings GetActiveProvider()
        {
            return Providers?.FirstOrDefault(provider =>
                string.Equals(provider.Id, ActiveProviderId,
                    StringComparison.OrdinalIgnoreCase));
        }

        public NlpProviderConfiguration Clone()
        {
            return new NlpProviderConfiguration
            {
                ActiveProviderId = ActiveProviderId,
                Providers = Providers == null
                    ? new List<NlpProviderSettings>()
                    : Providers.Select(provider => provider.Clone())
                        .ToList()
            };
        }
    }
}
