using ALPS_Visio_AddIn_rewrite.Importing;
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace ALPS_Visio_AddIn_rewrite
{
    /// <summary>
    /// UI-facing entry point for importing OWL ALPS files.
    /// </summary>
    public class OWLImporter
    {
        private readonly Lazy<OwlImportService> importService;
        private readonly Action<string, string> errorPresenter;

        /// <summary>
        /// Retained for source compatibility. The Ribbon owns a separate importer.
        /// </summary>
        [Obsolete("Create and retain an OWLImporter instance instead.")]
        public static readonly OWLImporter Instance = new OWLImporter();

        public OWLImporter()
            : this(new Lazy<OwlImportService>(OwlImportComposition.CreateDefault), ShowError)
        {
        }

        internal OWLImporter(Lazy<OwlImportService> importService,
            Action<string, string> errorPresenter)
        {
            this.importService = importService
                ?? throw new ArgumentNullException(nameof(importService));
            this.errorPresenter = errorPresenter
                ?? throw new ArgumentNullException(nameof(errorPresenter));
        }

        public void Parse(string fileName)
        {
            try
            {
                importService.Value.Import(fileName);
            }
            catch (Exception exception)
            {
                Debug.WriteLine("OWL import failed:");
                Debug.WriteLine(exception.ToString());
                errorPresenter("The OWL model could not be imported.\n\n"
                    + exception.Message, "ALPS/PASS Import");
            }
        }

        private static void ShowError(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
