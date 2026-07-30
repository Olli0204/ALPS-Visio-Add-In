using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ALPS_Visio_AddIn_rewrite.Verification
{
    internal sealed class VerificationResultsForm : Form
    {
        private readonly VerificationReport report;
        private readonly DataGridView grid;

        public VerificationResultsForm(VerificationReport report)
        {
            this.report = report
                ?? throw new ArgumentNullException(nameof(report));

            Text = "ALPS Verification – Ergebnisse";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(850, 480);
            Size = new Size(1180, 680);

            Label summary = new Label
            {
                Dock = DockStyle.Top,
                Height = 52,
                Padding = new Padding(10, 9, 10, 5),
                Text = CreateSummary(),
                BackColor = report.PassedSupportedChecks
                    ? Color.Honeydew
                    : Color.MistyRose
            };

            Label files = new Label
            {
                Dock = DockStyle.Top,
                Height = 46,
                Padding = new Padding(10, 5, 10, 5),
                Text = "Spezifikation: "
                    + Path.GetFileName(report.SpecificationFile)
                    + "\r\nImplementierung: "
                    + Path.GetFileName(report.ImplementationFile)
            };

            grid = CreateGrid();
            grid.Dock = DockStyle.Fill;
            foreach (VerificationFinding finding in report.Findings
                .OrderByDescending(item => item.Severity))
            {
                AddRow(finding);
            }

            FlowLayoutPanel buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 45,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(8)
            };
            Button closeButton = new Button
            {
                Text = "Schließen",
                DialogResult = DialogResult.OK,
                AutoSize = true
            };
            Button copyButton = new Button
            {
                Text = "Ergebnisse kopieren",
                AutoSize = true
            };
            copyButton.Click += CopyResults;
            buttons.Controls.Add(closeButton);
            buttons.Controls.Add(copyButton);

            Controls.Add(grid);
            Controls.Add(files);
            Controls.Add(summary);
            Controls.Add(buttons);
            AcceptButton = closeButton;
        }

        private string CreateSummary()
        {
            string status = report.PassedSupportedChecks
                ? "Die unterstützten Prüfungen wurden bestanden."
                : "Die Implementierung verletzt mindestens eine "
                    + "unterstützte Prüfregel.";
            return status + "\r\n" + report.CheckedRuleCount
                + " Regeln geprüft; " + report.ErrorCount + " Fehler; "
                + report.WarningCount + " Warnungen.";
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
            result.Columns.Add("Severity", "Schweregrad");
            result.Columns.Add("Category", "Bereich");
            result.Columns.Add("Code", "Regel");
            result.Columns.Add("Specification", "Spezifikationselement");
            result.Columns.Add("Implementation", "Implementierungselement");
            result.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Message",
                HeaderText = "Befund",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 320,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    WrapMode = DataGridViewTriState.True
                }
            });
            result.AutoResizeColumns(
                DataGridViewAutoSizeColumnsMode.AllCells);
            return result;
        }

        private void AddRow(VerificationFinding finding)
        {
            int rowIndex = grid.Rows.Add(
                SeverityText(finding.Severity),
                finding.Category ?? string.Empty,
                finding.Code ?? string.Empty,
                finding.SpecificationElement ?? string.Empty,
                finding.ImplementationElement ?? string.Empty,
                finding.Message ?? string.Empty);
            grid.Rows[rowIndex].DefaultCellStyle.BackColor =
                SeverityColor(finding.Severity);
        }

        private void CopyResults(object sender, EventArgs e)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine(CreateSummary());
            text.AppendLine("Spezifikation\t"
                + report.SpecificationFile);
            text.AppendLine("Implementierung\t"
                + report.ImplementationFile);
            text.AppendLine();
            text.AppendLine(
                "Schweregrad\tBereich\tRegel\tSpezifikationselement"
                + "\tImplementierungselement\tBefund");

            foreach (VerificationFinding finding in report.Findings)
            {
                text.Append(SeverityText(finding.Severity)).Append('\t')
                    .Append(Clean(finding.Category)).Append('\t')
                    .Append(Clean(finding.Code)).Append('\t')
                    .Append(Clean(finding.SpecificationElement))
                    .Append('\t')
                    .Append(Clean(finding.ImplementationElement))
                    .Append('\t')
                    .Append(Clean(finding.Message))
                    .AppendLine();
            }

            Clipboard.SetText(text.ToString());
        }

        private static string SeverityText(
            VerificationSeverity severity)
        {
            switch (severity)
            {
                case VerificationSeverity.Error:
                    return "Fehler";
                case VerificationSeverity.Warning:
                    return "Warnung";
                default:
                    return "Information";
            }
        }

        private static Color SeverityColor(
            VerificationSeverity severity)
        {
            switch (severity)
            {
                case VerificationSeverity.Error:
                    return Color.MistyRose;
                case VerificationSeverity.Warning:
                    return Color.LemonChiffon;
                default:
                    return Color.WhiteSmoke;
            }
        }

        private static string Clean(string value)
        {
            return (value ?? string.Empty)
                .Replace('\t', ' ')
                .Replace("\r", " ")
                .Replace("\n", " | ");
        }
    }
}
