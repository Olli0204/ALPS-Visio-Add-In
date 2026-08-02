using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using ALPS_Visio_AddIn_rewrite.VisioInfrastructure;

namespace ALPS_Visio_AddIn_rewrite.Verification
{
    internal sealed class AlpsVerificationController
    {
        private const string FileFilter =
            "Ontology files (*.owl;*.rdf)|*.owl;*.rdf|"
            + "OWL files (*.owl)|*.owl|RDF files (*.rdf)|*.rdf";

        private readonly AlpsVerificationService service =
            new AlpsVerificationService();
        private readonly Func<string> exportCurrentModel;
        private bool isRunning;

        public AlpsVerificationController()
            : this(() =>
                new CurrentVisioModelOwlExporter().Export())
        {
        }

        internal AlpsVerificationController(
            Func<string> exportCurrentModel)
        {
            this.exportCurrentModel = exportCurrentModel
                ?? throw new ArgumentNullException(
                    nameof(exportCurrentModel));
        }

        public void Run()
        {
            Run(SelectFiles);
        }

        public void RunWithCurrentModelAsSpecification()
        {
            Run(() => SelectFilesWithCurrentModel(
                currentModelIsSpecification: true));
        }

        public void RunWithCurrentModelAsImplementation()
        {
            Run(() => SelectFilesWithCurrentModel(
                currentModelIsSpecification: false));
        }

        private void Run(Func<VerificationFiles?> fileProvider)
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

            isRunning = true;
            Cursor previousCursor = Cursor.Current;
            try
            {
                VerificationFiles? files = fileProvider();
                if (files == null)
                    return;

                Cursor.Current = Cursors.WaitCursor;
                VerificationReport report = service.Verify(
                    files.SpecificationFile,
                    files.ImplementationFile);
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

        private VerificationFiles? SelectFiles()
        {
            string? specificationFile = SelectFile(
                "Spezifikationsmodell auswählen", null);
            if (specificationFile == null)
                return null;

            string? implementationFile = SelectFile(
                "Implementierungsmodell auswählen",
                Path.GetDirectoryName(specificationFile));
            return implementationFile == null
                ? null
                : new VerificationFiles(
                    specificationFile, implementationFile);
        }

        private VerificationFiles? SelectFilesWithCurrentModel(
            bool currentModelIsSpecification)
        {
            string? otherFile = SelectFile(
                currentModelIsSpecification
                    ? "Implementierungsmodell auswählen"
                    : "Spezifikationsmodell auswählen",
                null);
            if (otherFile == null)
                return null;

            string currentModelFile = exportCurrentModel();
            return currentModelIsSpecification
                ? new VerificationFiles(currentModelFile, otherFile)
                : new VerificationFiles(otherFile, currentModelFile);
        }

        private static string? SelectFile(string title,
            string? initialDirectory)
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

        private sealed class VerificationFiles
        {
            public VerificationFiles(
                string specificationFile,
                string implementationFile)
            {
                SpecificationFile = specificationFile;
                ImplementationFile = implementationFile;
            }

            public string SpecificationFile { get; }

            public string ImplementationFile { get; }
        }
    }
}
