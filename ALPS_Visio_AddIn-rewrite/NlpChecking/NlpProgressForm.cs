using System;
using System.Drawing;
using System.Windows.Forms;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    internal sealed class NlpProgressForm : Form
    {
        private readonly Label statusLabel;
        private readonly ProgressBar progressBar;

        public NlpProgressForm()
        {
            Text = "NLP PASS Checking";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ControlBox = false;
            ClientSize = new Size(430, 85);

            statusLabel = new Label
            {
                AutoSize = false,
                Location = new Point(14, 12),
                Size = new Size(400, 20),
                Text = "Preparing model check…"
            };
            progressBar = new ProgressBar
            {
                Location = new Point(14, 42),
                Size = new Size(400, 22),
                Minimum = 0,
                Maximum = 1
            };

            Controls.Add(statusLabel);
            Controls.Add(progressBar);
        }

        public void SetProgress(int completed, int total)
        {
            total = Math.Max(1, total);
            progressBar.Maximum = total;
            progressBar.Value = Math.Min(Math.Max(0, completed), total);
            statusLabel.Text = "Checking model labels… "
                + completed + " / " + total;
            Application.DoEvents();
        }
    }
}
