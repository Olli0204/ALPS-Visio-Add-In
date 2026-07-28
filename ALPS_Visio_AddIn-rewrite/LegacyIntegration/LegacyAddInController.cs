using System;
using VisioAddIn;
using VisioAddIn.Snapping;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.LegacyIntegration
{
    /// <summary>
    /// Isolates the supported legacy snapping and layer-explorer subsystem
    /// from the VSTO host lifecycle.
    /// </summary>
    internal sealed class LegacyAddInController : IDisposable
    {
        private readonly ThisAddIn addIn;
        private readonly Visio.Application application;
        private Visio.Document activeDocument;
        private WindowDirectory layerExplorer;
        private bool started;

        public LegacyAddInController(ThisAddIn addIn, Visio.Application application)
        {
            this.addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ModelController ModelController { get; private set; }

        public void Start()
        {
            if (started) return;

            ModelController = CreateModelController();
            application.DocumentCreated += OnDocumentCreated;
            application.PageAdded += OnPageAdded;
            application.WindowActivated += OnWindowActivated;
            application.DocumentOpened += OnDocumentOpened;
            activeDocument = application.ActiveDocument;
            started = true;
        }

        public void Update()
        {
            EnsureStarted();
            ModelController.updateWholeController(application.ActiveDocument.Pages);
            layerExplorer.displayTreeView(ModelController.getTreeView());
        }

        public void ExtendsChanged(SIDPage extends, SIDPage changedPage)
        {
            EnsureStarted();
            ModelController.updateWholeController(application.ActiveDocument.Pages);
            ModelController.updateBackground(extends, changedPage);
            layerExplorer.displayTreeView(ModelController.getTreeView());
        }

        public void ShowDirectory()
        {
            EnsureStarted();
            ModelController.updateWholeController(application.ActiveDocument.Pages);

            AnchorBarsUsage anchorBar = new AnchorBarsUsage(addIn, ModelController);
            layerExplorer = anchorBar.CreateAnchorBar(application);
            layerExplorer.displayTreeView(ModelController.getTreeView());
        }

        public void RefreshLayerExplorer()
        {
            layerExplorer?.displayTreeView(ModelController.getTreeView());
        }

        public void Dispose()
        {
            if (!started) return;

            application.DocumentCreated -= OnDocumentCreated;
            application.PageAdded -= OnPageAdded;
            application.WindowActivated -= OnWindowActivated;
            application.DocumentOpened -= OnDocumentOpened;
            started = false;
        }

        private void OnWindowActivated(Visio.Window window)
        {
            if (application.ActiveDocument == null) return;
            if (activeDocument != null
                && string.Equals(activeDocument.FullName, application.ActiveDocument.FullName,
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            activeDocument = application.ActiveDocument;
            Reset();
        }

        private void OnPageAdded(Visio.Page page)
        {
            ModelController.pageAdded(page);
            RefreshLayerExplorer();
        }

        private void OnDocumentOpened(Visio.IVDocument document)
        {
            activeDocument = application.ActiveDocument;
            Reset();
        }

        private void OnDocumentCreated(Visio.IVDocument document)
        {
            activeDocument = application.ActiveDocument;
            Reset();
        }

        private void Reset()
        {
            ModelController = CreateModelController();
            ModelController.updateWholeController(application.ActiveDocument.Pages);
            RefreshLayerExplorer();
        }

        private ModelController CreateModelController()
        {
            return new ModelController(addIn);
        }

        private void EnsureStarted()
        {
            if (!started)
                throw new InvalidOperationException("The legacy add-in controller has not been started.");
        }
    }
}
