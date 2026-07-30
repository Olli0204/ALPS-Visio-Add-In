using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    internal sealed class NlpCheckingController
    {
        private readonly Func<Visio.Document> documentProvider;
        private readonly NlpNameClassifier classifier =
            new NlpNameClassifier();
        private readonly NlpProviderStore providerStore =
            new NlpProviderStore();
        private readonly NlpSuggestionClient suggestionClient =
            new NlpSuggestionClient();
        private bool isRunning;

        public NlpCheckingController(
            Func<Visio.Document> documentProvider)
        {
            this.documentProvider = documentProvider
                ?? throw new ArgumentNullException(nameof(documentProvider));
        }

        public async Task RunAsync()
        {
            if (isRunning)
            {
                MessageBox.Show(
                    "A model naming check is already running.",
                    "NLP PASS Checking",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            Visio.Document document = documentProvider();
            if (document == null)
            {
                MessageBox.Show(
                    "No Visio drawing is active.",
                    "NLP PASS Checking",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            isRunning = true;
            NlpProgressForm progress = null;
            try
            {
                IList<NlpShapeCandidate> candidates =
                    NlpShapeCollector.Collect(document);
                if (candidates.Count == 0)
                {
                    MessageBox.Show(
                        "No supported PASS shapes were found in the active drawing.",
                        "NLP PASS Checking",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                progress = new NlpProgressForm();
                progress.Show();
                progress.SetProgress(0, candidates.Count);

                classifier.EnsureTrained();
                NlpProviderConfiguration providerConfiguration =
                    providerStore.Load();
                NlpProviderSettings activeProvider =
                    providerConfiguration.GetActiveProvider()?.Clone();
                List<NlpCheckResult> results =
                    new List<NlpCheckResult>(candidates.Count);

                for (int index = 0; index < candidates.Count; index++)
                {
                    NlpShapeCandidate candidate = candidates[index];
                    NlpNamePrediction prediction = classifier.Predict(
                        candidate.Label, candidate.ShapeType);
                    string suggestions = string.Empty;

                    if (!prediction.IsValid
                        && activeProvider?.CanRequestSuggestions == true)
                    {
                        try
                        {
                            suggestions =
                                await suggestionClient.SuggestAsync(
                                    candidate, activeProvider);
                        }
                        catch (Exception exception)
                        {
                            suggestions =
                                "Suggestions unavailable: "
                                + exception.Message;
                        }
                    }

                    results.Add(new NlpCheckResult
                    {
                        Candidate = candidate,
                        IsValid = prediction.IsValid,
                        Probability = prediction.Probability,
                        Suggestions = suggestions
                    });
                    progress.SetProgress(index + 1, candidates.Count);
                }

                progress.Close();
                progress = null;
                using (NlpResultsForm resultsForm =
                    new NlpResultsForm(results))
                {
                    resultsForm.ShowDialog();
                }
            }
            catch (Exception exception)
            {
                System.Diagnostics.Debug.WriteLine(
                    "NLP model naming check failed: " + exception);
                MessageBox.Show(
                    "The model naming check failed.\r\n\r\n"
                        + exception.Message,
                    "NLP PASS Checking",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (progress != null)
                    progress.Close();
                isRunning = false;
            }
        }

        public void Retrain()
        {
            try
            {
                int count = classifier.Retrain();
                MessageBox.Show(
                    "The naming classifier was retrained from "
                        + count + " bundled examples.",
                    "NLP PASS Checking",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                System.Diagnostics.Debug.WriteLine(
                    "NLP classifier retraining failed: " + exception);
                MessageBox.Show(
                    "The naming classifier could not be retrained.\r\n\r\n"
                        + exception.Message,
                    "NLP PASS Checking",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void ConfigureProviders()
        {
            try
            {
                NlpProviderConfiguration configuration =
                    providerStore.Load();
                using (NlpProviderSettingsDialog dialog =
                    new NlpProviderSettingsDialog(
                        configuration, suggestionClient))
                {
                    if (dialog.ShowDialog() != DialogResult.OK)
                        return;

                    providerStore.Save(dialog.Configuration);
                    NlpProviderSettings active =
                        dialog.Configuration.GetActiveProvider();
                    MessageBox.Show(
                        "Provider settings were stored securely.\r\n\r\n"
                            + "Active provider: "
                            + (active?.DisplayName ?? "None")
                            + "\r\nModel: "
                            + (string.IsNullOrWhiteSpace(active?.Model)
                                ? "Not selected"
                                : active.Model),
                        "NLP PASS Checking",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception exception)
            {
                System.Diagnostics.Debug.WriteLine(
                    "NLP provider storage failed: " + exception);
                MessageBox.Show(
                    "The provider settings could not be stored.\r\n\r\n"
                        + exception.Message,
                    "NLP PASS Checking",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
