using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ALPS_Visio_AddIn_rewrite.Verification
{
    internal sealed class AlpsVerificationController
    {
        private const string FileFilter =
            "Ontology files (*.owl;*.rdf)|*.owl;*.rdf|"
            + "OWL files (*.owl)|*.owl|RDF files (*.rdf)|*.rdf";

        private readonly AlpsVerificationService service =
            new AlpsVerificationService();
        private bool isRunning;

        public void Run()
        {
            if (isRunning)
            {
                MessageBox.Show(
                    "Eine ALPS-Verifikation wird bereits ausgeführt.",
                    "ALPS Verification",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string specificationFile = SelectFile(
                "Spezifikationsmodell auswählen", null);
            if (specificationFile == null)
                return;

            string implementationFile = SelectFile(
                "Implementierungsmodell auswählen",
                Path.GetDirectoryName(specificationFile));
            if (implementationFile == null)
                return;

            isRunning = true;
            Cursor previousCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                VerificationReport report = service.Verify(
                    specificationFile, implementationFile);
                Cursor.Current = previousCursor;
                using (VerificationResultsForm form =
                    new VerificationResultsForm(report))
                {
                    form.ShowDialog();
                }
            }
            catch (Exception exception)
            {
                Debug.WriteLine(
                    "ALPS verification failed: " + exception);
                MessageBox.Show(
                    "Die ALPS-Modelle konnten nicht verifiziert werden."
                        + "\r\n\r\n" + exception.Message,
                    "ALPS Verification",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = previousCursor;
                isRunning = false;
            }
        }

        private static string SelectFile(string title,
            string initialDirectory)
        {
            using (OpenFileDialog dialog = new OpenFileDialog
            {
                Title = title,
                Filter = FileFilter,
                CheckFileExists = true,
                Multiselect = false,
                InitialDirectory = initialDirectory ?? string.Empty
            })
            {
                return dialog.ShowDialog() == DialogResult.OK
                    ? dialog.FileName
                    : null;
            }
        }
    }
}
