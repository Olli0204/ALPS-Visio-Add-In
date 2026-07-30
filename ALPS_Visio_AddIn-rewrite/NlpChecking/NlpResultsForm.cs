using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    internal sealed class NlpResultsForm : Form
    {
        private readonly DataGridView grid;
        private readonly IList<NlpCheckResult> results;

        public NlpResultsForm(IList<NlpCheckResult> results)
        {
            this.results = results
                ?? throw new ArgumentNullException(nameof(results));

            Text = "NLP PASS Checking – Results";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(780, 430);
            Size = new Size(1050, 620);

            int invalidCount = results.Count(result => !result.IsValid);
            Label summary = new Label
            {
                Dock = DockStyle.Top,
                Height = 42,
                Padding = new Padding(10, 12, 10, 4),
                Text = results.Count + " labels checked; "
                    + invalidCount + " require review."
            };

            grid = CreateGrid();
            grid.Dock = DockStyle.Fill;
            foreach (NlpCheckResult result in results)
                AddRow(result);

            FlowLayoutPanel buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 45,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(8)
            };
            Button closeButton = new Button
            {
                Text = "Close",
                DialogResult = DialogResult.OK,
                AutoSize = true
            };
            Button copyButton = new Button
            {
                Text = "Copy results",
                AutoSize = true
            };
            copyButton.Click += CopyResults;
            buttons.Controls.Add(closeButton);
            buttons.Controls.Add(copyButton);

            Controls.Add(grid);
            Controls.Add(summary);
            Controls.Add(buttons);
            AcceptButton = closeButton;
        }

        private static DataGridView CreateGrid()
        {
            DataGridView result = new DataGridView
            {
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToOrderColumns = true,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            result.Columns.Add("Page", "Page");
            result.Columns.Add("Shape", "Shape ID");
            result.Columns.Add("Type", "Shape type");
            result.Columns.Add("Label", "Current label");
            result.Columns.Add("Result", "Result");
            result.Columns.Add("Confidence", "Confidence");
            DataGridViewTextBoxColumn suggestions =
                new DataGridViewTextBoxColumn
                {
                    Name = "Suggestions",
                    HeaderText = "Suggestions",
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                    MinimumWidth = 240,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        WrapMode = DataGridViewTriState.True
                    }
                };
            result.Columns.Add(suggestions);
            result.AutoResizeColumns(
                DataGridViewAutoSizeColumnsMode.AllCells);
            return result;
        }

        private void AddRow(NlpCheckResult result)
        {
            int rowIndex = grid.Rows.Add(
                result.Candidate.PageName,
                result.Candidate.ShapeId,
                result.Candidate.ShapeType,
                result.Candidate.Label,
                result.IsValid ? "Valid" : "Review",
                FormatConfidence(result),
                result.Suggestions ?? string.Empty);

            grid.Rows[rowIndex].DefaultCellStyle.BackColor =
                result.IsValid
                    ? Color.Honeydew
                    : Color.MistyRose;
        }

        private void CopyResults(object sender, EventArgs e)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine(
                "Page\tShape ID\tShape type\tLabel\tResult\tConfidence\tSuggestions");
            foreach (NlpCheckResult result in results)
            {
                text.Append(result.Candidate.PageName).Append('\t')
                    .Append(result.Candidate.ShapeId).Append('\t')
                    .Append(result.Candidate.ShapeType).Append('\t')
                    .Append(Clean(result.Candidate.Label)).Append('\t')
                    .Append(result.IsValid ? "Valid" : "Review").Append('\t')
                    .Append(FormatConfidence(result)).Append('\t')
                    .Append(Clean(result.Suggestions))
                    .AppendLine();
            }

            Clipboard.SetText(text.ToString());
        }

        private static string Clean(string value)
        {
            return (value ?? string.Empty)
                .Replace('\t', ' ')
                .Replace("\r", " ")
                .Replace("\n", " | ");
        }

        private static string FormatConfidence(NlpCheckResult result)
        {
            float probability = result.IsValid
                ? result.Probability
                : 1.0f - result.Probability;
            return probability.ToString("P0");
        }
    }
}
