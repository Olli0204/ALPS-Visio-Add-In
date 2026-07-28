using Microsoft.Office.Interop.Visio;
using System.Windows.Forms;
using VisioAddIn;
using VisioAddIn.Snapping;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// Entrypoint of this AddIn.
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            prepStuff();
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }

        #endregion

        #region Model Explorer and snapping integration

        private void prepStuff()
        {
            modelManager = new ModelController(this);

            // Add triggers for methods to be called when an updateWholeController in visio occurs
            Application.DocumentCreated += Application_DocumentCreated;
            Application.PageAdded += Application_PageAdded;
            Application.WindowActivated += Application_WindowActivated;
            Application.DocumentOpened += Application_DocumentOpened;

            // Set the current active document
            Visio.Document document = Application.ActiveDocument;
            if (IsDrawingDocument(document))
                activeDoc = document;
        }

        /// <summary>
        /// The active Visio document this Add-In operates in
        /// </summary>
        private Visio.Document activeDoc;


        /// <summary>
        /// reference to Directory where TreeView etc is displayed.
        /// </summary>
        private WindowDirectory layerExplorer;


        /// <summary>
        /// reference to the ModelManager where the data is maintained
        /// </summary>
        private ModelController modelManager;

        /// <summary>
        /// Called when the active window in the document changes.
        /// Checks whether the active document is still the same or not.
        /// </summary>
        /// <param name="window">The active window, not used by this function</param>
        private void Application_WindowActivated(Window window)
        {
            Visio.Document document = Application.ActiveDocument;
            if (!IsDrawingDocument(document)) return;

            if (IsSameDocument(activeDoc, document)) return;
            activeDoc = document;
            reset();
        }

        /// <summary>
        /// Called when a Page was added. Determines to which model the Page belongs to
        /// </summary>
        private void Application_PageAdded(Page page)
        {
            if (page == null || !IsDrawingDocument(page.Document))
                return;

            //let the model manager determine to what model the new Page belongs to
            modelManager.pageAdded(page);

            refreshLayerExplorerTreeView();
        }

        private void Application_DocumentOpened(IVDocument doc)
        {
            Visio.Document document = doc as Visio.Document;
            if (!IsDrawingDocument(document)) return;

            activeDoc = document;
            reset();
        }

        private void Application_DocumentCreated(IVDocument doc)
        {
            Visio.Document document = doc as Visio.Document;
            if (!IsDrawingDocument(document)) return;

            activeDoc = document;
            reset();
        }

        internal void updateClicked()
        {
            Visio.Document document = GetDrawingDocument();
            if (document == null) return;

            modelManager.updateWholeController(document.Pages);
            layerExplorer.displayTreeView(modelManager.getTreeView());
        }

        internal ModelController getModelController()
        {
            return modelManager;
        }

        internal void RefreshModelFromDrawing()
        {
            reset();
        }

        internal void extendsChanged(SIDPage extends, SIDPage changedPage)
        {
            Visio.Document document = GetDrawingDocument();
            if (document == null) return;

            modelManager.updateWholeController(document.Pages);
            //if (changedPage.)
            modelManager.updateBackground(extends, changedPage);
            layerExplorer.displayTreeView(modelManager.getTreeView());
        }

        internal void showDirectoryClicked()
        {
            Visio.Document document = GetDrawingDocument();
            if (document == null) return;

            modelManager.updateWholeController(document.Pages);

            //Methods are not used due to a problem with the setParent-Method regarding the anchor-bar
            AnchorBarsUsage ancBar = new AnchorBarsUsage(this, modelManager);
            layerExplorer = ancBar.CreateAnchorBar(Application);

            layerExplorer.displayTreeView(modelManager.getTreeView());
        }
        private void reset()
        {
            Visio.Document document = GetDrawingDocument();
            if (document == null) return;

            this.modelManager = new ModelController(this);
            modelManager.updateWholeController(document.Pages);
            layerExplorer?.displayTreeView(modelManager.getTreeView());
        }

        /// <summary>
        /// Returns the most recently active drawing. Stencil windows can become
        /// ActiveDocument while they are opened and must never replace the model.
        /// </summary>
        internal Visio.Document GetDrawingDocument()
        {
            Visio.Document document = Application.ActiveDocument;
            if (IsDrawingDocument(document))
                activeDoc = document;

            return IsDrawingDocument(activeDoc) ? activeDoc : null;
        }

        private static bool IsDrawingDocument(IVDocument document)
        {
            if (document == null) return false;

            try
            {
                return document.Type == Visio.VisDocumentTypes.visTypeDrawing;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // A closing COM document is no longer a valid model source.
                return false;
            }
        }

        private static bool IsSameDocument(Visio.Document first,
            Visio.Document second)
        {
            if (first == null || second == null) return false;

            try
            {
                return string.Equals(first.FullName, second.FullName,
                    System.StringComparison.OrdinalIgnoreCase);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // A closing document cannot be the active model.
                return false;
            }
        }

        public void refreshLayerExplorerTreeView()
        {
            layerExplorer?.displayTreeView(modelManager.getTreeView());
        }

        #endregion
    }
}
