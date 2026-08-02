using ALPS_Visio_AddIn_rewrite.NlpChecking;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ALPS_Visio_AddIn_rewrite.Tests.NlpChecking
{
    [TestClass]
    public sealed class NlpProviderModelsTests
    {
        [TestMethod]
        public void CanRequestSuggestions_BuiltInProviderRequiresApiKey()
        {
            NlpProviderSettings provider = new NlpProviderSettings
            {
                BaseUrl = "https://provider.example/v1",
                Model = "model-a",
                IsBuiltIn = true
            };

            Assert.IsFalse(provider.CanRequestSuggestions);
            provider.ApiKey = "secret";
            Assert.IsTrue(provider.CanRequestSuggestions);
        }

        [TestMethod]
        public void CanRequestSuggestions_CustomProviderAllowsKeylessEndpoint()
        {
            NlpProviderSettings provider = new NlpProviderSettings
            {
                BaseUrl = "http://localhost:8080/v1",
                Model = "local-model",
                IsBuiltIn = false
            };

            Assert.IsTrue(provider.CanRequestSuggestions);
        }

        [TestMethod]
        public void Clone_Configuration_CreatesIndependentProviderLists()
        {
            NlpProviderConfiguration original =
                new NlpProviderConfiguration
                {
                    ActiveProviderId = "openai",
                    Providers = new List<NlpProviderSettings>
                    {
                        new NlpProviderSettings
                        {
                            Id = "openai",
                            DisplayName = "OpenAI",
                            KnownModels = new List<string> { "model-a" }
                        }
                    }
                };

            NlpProviderConfiguration clone = original.Clone();
            clone.Providers[0].DisplayName = "Changed";
            clone.Providers[0].KnownModels.Add("model-b");

            Assert.AreEqual("OpenAI", original.Providers[0].DisplayName);
            CollectionAssert.AreEqual(
                new[] { "model-a" },
                original.Providers[0].KnownModels);
            Assert.AreEqual("openai", clone.GetActiveProvider().Id);
        }
    }
}
