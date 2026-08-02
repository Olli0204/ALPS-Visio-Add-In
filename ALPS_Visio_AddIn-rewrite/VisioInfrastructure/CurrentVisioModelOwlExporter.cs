using System;
using System.IO;
using System.Runtime.InteropServices;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Runs the OWL exporter embedded in the interactive SID stencil for the
    /// currently active Visio drawing.
    /// </summary>
    internal sealed class CurrentVisioModelOwlExporter
    {
        internal const string ExportMacro =
            "ALPS_RDFOWLExporter.createProcessRDFOWL";
        internal const string EnableMinimal2DExport =
            "ALPS_RDFOWLExporter."
            + "includeMinimal2DVisualisationDataInOWL = True";

        private readonly Func<Visio.Application> applicationProvider;
        private readonly Func<Visio.Document> drawingProvider;
        private readonly Action openInteractiveStencils;
        private readonly Func<Visio.Document> sidStencilProvider;

        public CurrentVisioModelOwlExporter()
            : this(
                () => Globals.ThisAddIn.Application,
                () => Globals.ThisAddIn.GetDrawingDocument(),
                VisioHelper.OpenInteractiveStencils,
                () => VisioHelper.openStencil(
                    VisioHelper.VisioStencils.SID_STENCIL))
        {
        }

        internal CurrentVisioModelOwlExporter(
            Func<Visio.Application> applicationProvider,
            Func<Visio.Document> drawingProvider,
            Action openInteractiveStencils,
            Func<Visio.Document> sidStencilProvider)
        {
            this.applicationProvider = applicationProvider
                ?? throw new ArgumentNullException(
                    nameof(applicationProvider));
            this.drawingProvider = drawingProvider
                ?? throw new ArgumentNullException(nameof(drawingProvider));
            this.openInteractiveStencils = openInteractiveStencils
                ?? throw new ArgumentNullException(
                    nameof(openInteractiveStencils));
            this.sidStencilProvider = sidStencilProvider
                ?? throw new ArgumentNullException(
                    nameof(sidStencilProvider));
        }

        /// <summary>
        /// Exports the complete model represented by the active drawing and
        /// returns the OWL file written by the SID stencil macro.
        /// </summary>
        public string Export()
        {
            Visio.Application application = applicationProvider();
            Visio.Document drawing = drawingProvider();
            if (application == null || drawing == null)
            {
                throw new InvalidOperationException(
                    "Es ist kein Visio-Zeichnungsmodell aktiv.");
            }

            string outputFilePath = GetOutputFilePath(
                drawing.Path, drawing.Name);
            Visio.Window drawingWindow = GetActiveDrawingWindow(
                application, drawing);

            try
            {
                // The SID stencil owns ALPS_RDFOWLExporter. Opening it may
                // briefly activate the stencil window, so reactivate the
                // drawing before executing the macro: the VBA implementation
                // deliberately reads Visio.ActiveDocument.
                openInteractiveStencils();
                drawingWindow.Activate();

                if (!IsSameDocument(application.ActiveDocument, drawing))
                {
                    throw new InvalidOperationException(
                        "Das aktuelle Visio-Zeichnungsmodell konnte vor dem "
                        + "OWL-Export nicht aktiviert werden.");
                }

                Visio.Document sidStencil = sidStencilProvider();
                if (sidStencil == null)
                {
                    throw new InvalidOperationException(
                        "Der makrofähige ALPS-SID-Stencil konnte nicht "
                        + "geöffnet werden.");
                }

                sidStencil.ExecuteLine(EnableMinimal2DExport);
                sidStencil.ExecuteLine(ExportMacro);
            }
            catch (COMException exception)
            {
                throw new InvalidOperationException(
                    "Der OWL-Export des SID-Stencils konnte nicht "
                    + "ausgeführt werden. Prüfen Sie die Visio-"
                    + "Makroeinstellungen und ob der ALPS-SID-Stencil "
                    + "vertrauenswürdig geöffnet wurde.\r\n\r\n"
                    + exception.Message,
                    exception);
            }

            FileInfo outputFile = new FileInfo(outputFilePath);
            if (!outputFile.Exists || outputFile.Length == 0)
            {
                throw new InvalidOperationException(
                    "Das OWL-Exportmakro wurde ausgeführt, hat aber keine "
                    + "lesbare Datei erzeugt: " + outputFilePath);
            }

            return outputFile.FullName;
        }

        internal static string GetOutputFilePath(
            string documentPath, string documentName)
        {
            if (string.IsNullOrWhiteSpace(documentPath))
            {
                throw new InvalidOperationException(
                    "Das aktuelle Visio-Modell wurde noch nicht gespeichert. "
                    + "Speichern Sie die Zeichnung zuerst, damit der "
                    + "OWL-Exporter eine Zieldatei bestimmen kann.");
            }

            if (documentPath.IndexOf(
                    "http", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new InvalidOperationException(
                    "Das aktuelle Visio-Modell besitzt nur einen Webpfad. "
                    + "Speichern oder synchronisieren Sie es lokal, damit "
                    + "die vom OWL-Makro erzeugte Datei sicher an die "
                    + "Folgefunktion übergeben werden kann.");
            }

            if (string.IsNullOrWhiteSpace(documentName))
            {
                throw new InvalidOperationException(
                    "Der Name des aktuellen Visio-Modells ist nicht "
                    + "verfügbar.");
            }

            // Keep this transformation identical to createProcessRDFOWL in
            // the v1.0.1 SID stencil.
            string modelName = documentName
                .Replace(" ", "_")
                .Replace(".vsdx", string.Empty);
            return Path.Combine(documentPath, modelName + ".owl");
        }

        private static Visio.Window GetActiveDrawingWindow(
            Visio.Application application, Visio.Document drawing)
        {
            Visio.Window activeWindow = application.ActiveWindow;
            if (activeWindow == null
                || !IsSameDocument(application.ActiveDocument, drawing))
            {
                throw new InvalidOperationException(
                    "Aktivieren Sie zunächst ein Visio-Zeichnungsfenster "
                    + "mit dem zu exportierenden ALPS/PASS-Modell.");
            }

            return activeWindow;
        }

        private static bool IsSameDocument(
            Visio.Document first, Visio.Document second)
        {
            if (first == null || second == null)
                return false;

            try
            {
                return string.Equals(
                    first.FullName,
                    second.FullName,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (COMException)
            {
                return false;
            }
        }
    }
}
