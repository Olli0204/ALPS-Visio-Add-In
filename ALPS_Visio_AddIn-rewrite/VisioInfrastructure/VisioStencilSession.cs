using System;
using System.Runtime.InteropServices;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Opens stencil sets while preserving the active drawing window.
    /// </summary>
    internal sealed class VisioStencilSession
    {
        private readonly Func<Visio.Application> applicationProvider;
        private readonly Func<Visio.Document> drawingProvider;
        private readonly VisioStencilRepository stencilRepository;

        public VisioStencilSession(Func<Visio.Application> applicationProvider,
            Func<Visio.Document> drawingProvider,
            VisioStencilRepository stencilRepository)
        {
            this.applicationProvider = applicationProvider
                ?? throw new ArgumentNullException(nameof(applicationProvider));
            this.drawingProvider = drawingProvider
                ?? throw new ArgumentNullException(nameof(drawingProvider));
            this.stencilRepository = stencilRepository
                ?? throw new ArgumentNullException(nameof(stencilRepository));
        }

        public void OpenImportStencils()
        {
            Open(stencilRepository.OpenImportStencils, true);
        }

        public void OpenInteractiveStencils()
        {
            Open(stencilRepository.OpenInteractiveStencils, false);
        }

        private void Open(Action openStencils, bool createDrawingIfMissing)
        {
            Visio.Application application = applicationProvider();
            Visio.Window drawingWindow = null;
            try
            {
                Visio.Document drawing = drawingProvider();
                if (drawing == null && createDrawingIfMissing)
                    drawing = application.Documents.Add("");

                Visio.Document activeDocument = application.ActiveDocument;
                if (drawing != null
                    && activeDocument != null
                    && activeDocument.Type
                        == Visio.VisDocumentTypes.visTypeDrawing)
                {
                    drawingWindow = application.ActiveWindow;
                }

                openStencils();
            }
            finally
            {
                try
                {
                    drawingWindow?.Activate();
                }
                catch (COMException)
                {
                    // The drawing may have been closed while a stencil opened.
                }
            }
        }
    }
}
