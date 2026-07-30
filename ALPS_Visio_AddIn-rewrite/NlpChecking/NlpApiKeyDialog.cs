using System;
using System.Drawing;
using System.Windows.Forms;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    internal sealed class NlpApiKeyDialog : Form
    {
        private readonly TextBox apiKeyTextBox;

        public string ApiKey => apiKeyTextBox.Text.Trim();

        public NlpApiKeyDialog(string currentApiKey)
        {
            Text = "NLP PASS Checking – API Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 205);

            Label explanation = new Label
            {
                AutoSize = false,
                Location = new Point(16, 15),
                Size = new Size(488, 66),
                Text = "The model check runs locally. For invalid labels, "
                    + "an optional API request can generate suggestions.\r\n"
                    + "Only the shape type and label are sent to:\r\n"
                    + NlpSuggestionClient.Endpoint
            };

            Label keyLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 91),
                Text = "API key:"
            };

            apiKeyTextBox = new TextBox
            {
                Location = new Point(16, 112),
                Size = new Size(488, 23),
                UseSystemPasswordChar = true,
                Text = currentApiKey ?? string.Empty
            };

            Label storage = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Location = new Point(16, 141),
                Text = "Stored encrypted for the current Windows user. "
                    + "Leave blank to remove it."
            };

            Button okButton = new Button
            {
                Text = "Save",
                DialogResult = DialogResult.OK,
                Location = new Point(348, 169),
                Size = new Size(75, 25)
            };
            Button cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(429, 169),
                Size = new Size(75, 25)
            };

            Controls.Add(explanation);
            Controls.Add(keyLabel);
            Controls.Add(apiKeyTextBox);
            Controls.Add(storage);
            Controls.Add(okButton);
            Controls.Add(cancelButton);
            AcceptButton = okButton;
            CancelButton = cancelButton;
        }
    }
}
