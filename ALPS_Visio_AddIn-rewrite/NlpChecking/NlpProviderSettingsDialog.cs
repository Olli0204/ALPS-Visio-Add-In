using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    internal sealed class NlpProviderSettingsDialog : Form
    {
        private readonly NlpProviderConfiguration configuration;
        private readonly NlpSuggestionClient providerClient;
        private readonly ComboBox providerComboBox;
        private readonly TextBox nameTextBox;
        private readonly ComboBox protocolComboBox;
        private readonly TextBox baseUrlTextBox;
        private readonly TextBox apiKeyTextBox;
        private readonly ComboBox modelComboBox;
        private readonly CheckBox activeProviderCheckBox;
        private readonly CheckBox showApiKeyCheckBox;
        private readonly Button addProviderButton;
        private readonly Button removeProviderButton;
        private readonly Button refreshModelsButton;
        private readonly Button saveButton;
        private readonly Button cancelButton;
        private readonly Label statusLabel;
        private readonly Label privacyLabel;
        private NlpProviderSettings selectedProvider;
        private bool loadingEditor;
        private bool isBusy;

        public NlpProviderConfiguration Configuration =>
            configuration;

        public NlpProviderSettingsDialog(
            NlpProviderConfiguration configuration,
            NlpSuggestionClient providerClient)
        {
            this.configuration = configuration?.Clone()
                ?? throw new ArgumentNullException(nameof(configuration));
            this.providerClient = providerClient
                ?? throw new ArgumentNullException(nameof(providerClient));

            Text = "NLP PASS Checking – Provider Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(760, 490);

            providerComboBox = CreateComboBox(
                new Point(130, 17), new Size(380, 24),
                ComboBoxStyle.DropDownList);
            providerComboBox.SelectedIndexChanged +=
                ProviderSelectionChanged;
            addProviderButton = CreateButton(
                "Add provider", new Point(520, 16),
                new Size(105, 26));
            addProviderButton.Click += AddProvider;
            removeProviderButton = CreateButton(
                "Remove", new Point(635, 16),
                new Size(105, 26));
            removeProviderButton.Click += RemoveProvider;

            activeProviderCheckBox = new CheckBox
            {
                AutoSize = true,
                Location = new Point(130, 52),
                Text = "Use this provider for naming suggestions"
            };
            activeProviderCheckBox.CheckedChanged +=
                ActiveProviderChanged;

            nameTextBox = CreateTextBox(
                new Point(130, 87), new Size(610, 23));
            protocolComboBox = CreateComboBox(
                new Point(130, 124), new Size(280, 24),
                ComboBoxStyle.DropDownList);
            protocolComboBox.Items.Add(new ProtocolOption(
                NlpApiProtocol.OpenAiCompatible,
                "OpenAI-compatible"));
            protocolComboBox.Items.Add(new ProtocolOption(
                NlpApiProtocol.AnthropicCompatible,
                "Anthropic-compatible"));

            baseUrlTextBox = CreateTextBox(
                new Point(130, 161), new Size(610, 23));
            apiKeyTextBox = CreateTextBox(
                new Point(130, 198), new Size(520, 23));
            apiKeyTextBox.UseSystemPasswordChar = true;
            showApiKeyCheckBox = new CheckBox
            {
                AutoSize = true,
                Location = new Point(660, 201),
                Text = "Show"
            };
            showApiKeyCheckBox.CheckedChanged +=
                delegate
                {
                    apiKeyTextBox.UseSystemPasswordChar =
                        !showApiKeyCheckBox.Checked;
                };

            modelComboBox = CreateComboBox(
                new Point(130, 235), new Size(480, 24),
                ComboBoxStyle.DropDown);
            refreshModelsButton = CreateButton(
                "Load models", new Point(620, 234),
                new Size(120, 26));
            refreshModelsButton.Click += RefreshModels;

            statusLabel = new Label
            {
                AutoSize = false,
                Location = new Point(130, 270),
                Size = new Size(610, 38),
                ForeColor = SystemColors.GrayText,
                Text = "Enter an API key and load the available models."
            };

            GroupBox privacyGroup = new GroupBox
            {
                Location = new Point(15, 315),
                Size = new Size(725, 105),
                Text = "Data and security"
            };
            privacyLabel = new Label
            {
                AutoSize = false,
                Location = new Point(12, 22),
                Size = new Size(700, 70)
            };
            privacyGroup.Controls.Add(privacyLabel);

            saveButton = CreateButton(
                "Save", new Point(565, 445),
                new Size(80, 28));
            saveButton.Click += SaveConfiguration;
            cancelButton = CreateButton(
                "Cancel", new Point(660, 445),
                new Size(80, 28));
            cancelButton.DialogResult = DialogResult.Cancel;
            FormClosing += PreventCloseWhileBusy;

            AddLabel("Provider:", 15, 20);
            AddLabel("Name:", 15, 90);
            AddLabel("API protocol:", 15, 127);
            AddLabel("Base URL:", 15, 164);
            AddLabel("API key:", 15, 201);
            AddLabel("Model:", 15, 238);
            Controls.Add(providerComboBox);
            Controls.Add(addProviderButton);
            Controls.Add(removeProviderButton);
            Controls.Add(activeProviderCheckBox);
            Controls.Add(nameTextBox);
            Controls.Add(protocolComboBox);
            Controls.Add(baseUrlTextBox);
            Controls.Add(apiKeyTextBox);
            Controls.Add(showApiKeyCheckBox);
            Controls.Add(modelComboBox);
            Controls.Add(refreshModelsButton);
            Controls.Add(statusLabel);
            Controls.Add(privacyGroup);
            Controls.Add(saveButton);
            Controls.Add(cancelButton);
            AcceptButton = saveButton;
            CancelButton = cancelButton;

            BindProviders(this.configuration.ActiveProviderId);
        }

        private void BindProviders(string selectedId)
        {
            loadingEditor = true;
            providerComboBox.Items.Clear();
            foreach (NlpProviderSettings provider
                in configuration.Providers)
            {
                providerComboBox.Items.Add(provider);
            }

            NlpProviderSettings selection =
                configuration.Providers.FirstOrDefault(provider =>
                    string.Equals(provider.Id, selectedId,
                        StringComparison.OrdinalIgnoreCase))
                ?? configuration.Providers.FirstOrDefault();
            providerComboBox.SelectedItem = selection;
            selectedProvider = selection;
            loadingEditor = false;
            LoadProviderIntoEditor();
        }

        private void ProviderSelectionChanged(
            object sender, EventArgs e)
        {
            if (loadingEditor)
                return;

            SaveEditorToProvider();
            selectedProvider =
                providerComboBox.SelectedItem
                as NlpProviderSettings;
            LoadProviderIntoEditor();
        }

        private void LoadProviderIntoEditor()
        {
            if (selectedProvider == null)
                return;

            loadingEditor = true;
            nameTextBox.Text =
                selectedProvider.DisplayName ?? string.Empty;
            baseUrlTextBox.Text =
                selectedProvider.BaseUrl ?? string.Empty;
            apiKeyTextBox.Text =
                selectedProvider.ApiKey ?? string.Empty;
            SelectProtocol(selectedProvider.Protocol);
            PopulateModels(selectedProvider);
            activeProviderCheckBox.Checked = string.Equals(
                configuration.ActiveProviderId,
                selectedProvider.Id,
                StringComparison.OrdinalIgnoreCase);

            bool custom = !selectedProvider.IsBuiltIn;
            nameTextBox.ReadOnly = !custom;
            baseUrlTextBox.ReadOnly = !custom;
            protocolComboBox.Enabled = custom;
            removeProviderButton.Enabled = custom;
            privacyLabel.Text =
                "The full provider configuration is encrypted with "
                + "Windows DPAPI for the current user.\r\n"
                + "Model loading sends only an authenticated metadata "
                + "request. Naming suggestions send the shape type and "
                + "label to " + selectedProvider.BaseUrl + ". "
                + "The selected provider's privacy terms apply.";
            statusLabel.Text =
                selectedProvider.KnownModels?.Count > 0
                    ? selectedProvider.KnownModels.Count
                        + " model(s) available in the dropdown. "
                        + "Use Load models to refresh."
                    : "Enter an API key and load the available models.";
            loadingEditor = false;
        }

        private void SaveEditorToProvider()
        {
            if (selectedProvider == null || loadingEditor)
                return;

            if (!selectedProvider.IsBuiltIn)
            {
                selectedProvider.DisplayName =
                    nameTextBox.Text.Trim();
                selectedProvider.BaseUrl =
                    baseUrlTextBox.Text.Trim();
                selectedProvider.Protocol = SelectedProtocol;
            }

            selectedProvider.ApiKey = apiKeyTextBox.Text.Trim();
            selectedProvider.Model = modelComboBox.Text.Trim();
            if (activeProviderCheckBox.Checked)
            {
                configuration.ActiveProviderId =
                    selectedProvider.Id;
            }
        }

        private void PopulateModels(NlpProviderSettings provider)
        {
            modelComboBox.Items.Clear();
            IEnumerable<string> models =
                provider.KnownModels ?? new List<string>();
            foreach (string model in models
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value,
                    StringComparer.OrdinalIgnoreCase))
            {
                modelComboBox.Items.Add(model);
            }

            if (!string.IsNullOrWhiteSpace(provider.Model)
                && !modelComboBox.Items.Contains(provider.Model))
            {
                modelComboBox.Items.Add(provider.Model);
            }
            modelComboBox.Text = provider.Model ?? string.Empty;
        }

        private async void RefreshModels(
            object sender, EventArgs e)
        {
            SaveEditorToProvider();
            if (selectedProvider == null)
                return;

            if (!ValidateProvider(
                selectedProvider, requireCredentials: true))
                return;

            SetBusy(true);
            statusLabel.Text =
                "Loading models from " + selectedProvider.DisplayName
                + "…";
            try
            {
                IList<string> models =
                    await providerClient.GetModelsAsync(
                        selectedProvider.Clone());
                selectedProvider.KnownModels =
                    new List<string>(models);
                if (string.IsNullOrWhiteSpace(
                        selectedProvider.Model)
                    || !models.Contains(selectedProvider.Model,
                        StringComparer.OrdinalIgnoreCase))
                {
                    selectedProvider.Model = models[0];
                }

                loadingEditor = true;
                PopulateModels(selectedProvider);
                loadingEditor = false;
                statusLabel.Text = models.Count
                    + " model(s) loaded. Select one from the dropdown.";
            }
            catch (Exception exception)
            {
                System.Diagnostics.Debug.WriteLine(
                    "NLP provider model query failed: " + exception);
                statusLabel.Text =
                    "Models could not be loaded: "
                    + exception.Message;
                MessageBox.Show(
                    "The provider model list could not be loaded.\r\n\r\n"
                        + exception.Message,
                    "NLP PASS Checking",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private void AddProvider(object sender, EventArgs e)
        {
            SaveEditorToProvider();
            NlpProviderSettings provider =
                new NlpProviderSettings
                {
                    Id = "custom-"
                        + Guid.NewGuid().ToString("N"),
                    DisplayName = "Custom Provider",
                    BaseUrl = "https://",
                    Protocol =
                        NlpApiProtocol.OpenAiCompatible,
                    IsBuiltIn = false
                };
            configuration.Providers.Add(provider);
            BindProviders(provider.Id);
            nameTextBox.SelectAll();
            nameTextBox.Focus();
        }

        private void RemoveProvider(object sender, EventArgs e)
        {
            if (selectedProvider == null
                || selectedProvider.IsBuiltIn)
                return;

            DialogResult confirmation = MessageBox.Show(
                "Remove provider \"" + selectedProvider.DisplayName
                    + "\" and its stored API key?",
                "NLP PASS Checking",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (confirmation != DialogResult.Yes)
                return;

            string removedId = selectedProvider.Id;
            configuration.Providers.Remove(selectedProvider);
            if (string.Equals(configuration.ActiveProviderId,
                removedId, StringComparison.OrdinalIgnoreCase))
            {
                configuration.ActiveProviderId =
                    NlpProviderStore.UniGptId;
            }
            BindProviders(configuration.ActiveProviderId);
        }

        private void ActiveProviderChanged(
            object sender, EventArgs e)
        {
            if (loadingEditor || selectedProvider == null)
                return;

            if (activeProviderCheckBox.Checked)
            {
                configuration.ActiveProviderId =
                    selectedProvider.Id;
                statusLabel.Text = selectedProvider.DisplayName
                    + " is now the active suggestion provider.";
            }
            else if (string.Equals(
                configuration.ActiveProviderId,
                selectedProvider.Id,
                StringComparison.OrdinalIgnoreCase))
            {
                loadingEditor = true;
                activeProviderCheckBox.Checked = true;
                loadingEditor = false;
            }
        }

        private void SaveConfiguration(
            object sender, EventArgs e)
        {
            SaveEditorToProvider();
            foreach (NlpProviderSettings provider
                in configuration.Providers)
            {
                if (!ValidateProvider(
                    provider, requireCredentials: false))
                    return;
            }

            bool duplicateName = configuration.Providers
                .GroupBy(provider => provider.DisplayName.Trim(),
                    StringComparer.OrdinalIgnoreCase)
                .Any(group => group.Count() > 1);
            if (duplicateName)
            {
                MessageBox.Show(
                    "Provider names must be unique.",
                    "NLP PASS Checking",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private bool ValidateProvider(
            NlpProviderSettings provider,
            bool requireCredentials)
        {
            if (string.IsNullOrWhiteSpace(provider.DisplayName))
            {
                ShowValidationError(
                    "Every provider requires a name.");
                return false;
            }

            if (!Uri.TryCreate(provider.BaseUrl,
                    UriKind.Absolute, out Uri uri)
                || (uri.Scheme != Uri.UriSchemeHttps
                    && uri.Scheme != Uri.UriSchemeHttp))
            {
                ShowValidationError(
                    "Enter a valid HTTP or HTTPS base URL.");
                return false;
            }

            if (requireCredentials
                && provider.IsBuiltIn
                && string.IsNullOrWhiteSpace(provider.ApiKey))
            {
                ShowValidationError(
                    "Enter an API key for "
                    + provider.DisplayName + ".");
                return false;
            }

            return true;
        }

        private static void ShowValidationError(string message)
        {
            MessageBox.Show(
                message,
                "NLP PASS Checking",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private void SetBusy(bool busy)
        {
            isBusy = busy;
            providerComboBox.Enabled = !busy;
            addProviderButton.Enabled = !busy;
            removeProviderButton.Enabled =
                !busy && selectedProvider != null
                && !selectedProvider.IsBuiltIn;
            refreshModelsButton.Enabled = !busy;
            saveButton.Enabled = !busy;
            cancelButton.Enabled = !busy;
            UseWaitCursor = busy;
        }

        private void PreventCloseWhileBusy(
            object sender, FormClosingEventArgs e)
        {
            if (isBusy)
                e.Cancel = true;
        }

        private NlpApiProtocol SelectedProtocol
        {
            get
            {
                ProtocolOption option =
                    protocolComboBox.SelectedItem
                    as ProtocolOption;
                return option?.Protocol
                    ?? NlpApiProtocol.OpenAiCompatible;
            }
        }

        private void SelectProtocol(NlpApiProtocol protocol)
        {
            foreach (object item in protocolComboBox.Items)
            {
                ProtocolOption option = item as ProtocolOption;
                if (option?.Protocol == protocol)
                {
                    protocolComboBox.SelectedItem = option;
                    return;
                }
            }
            protocolComboBox.SelectedIndex = 0;
        }

        private void AddLabel(string text, int x, int y)
        {
            Controls.Add(new Label
            {
                AutoSize = true,
                Location = new Point(x, y),
                Text = text
            });
        }

        private static TextBox CreateTextBox(
            Point location, Size size)
        {
            return new TextBox
            {
                Location = location,
                Size = size
            };
        }

        private static ComboBox CreateComboBox(
            Point location, Size size,
            ComboBoxStyle style)
        {
            return new ComboBox
            {
                Location = location,
                Size = size,
                DropDownStyle = style
            };
        }

        private static Button CreateButton(
            string text, Point location, Size size)
        {
            return new Button
            {
                Text = text,
                Location = location,
                Size = size
            };
        }

        private sealed class ProtocolOption
        {
            public NlpApiProtocol Protocol { get; }
            private readonly string displayName;

            public ProtocolOption(
                NlpApiProtocol protocol, string displayName)
            {
                Protocol = protocol;
                this.displayName = displayName;
            }

            public override string ToString()
            {
                return displayName;
            }
        }
    }
}
