using ALPS_Visio_AddIn_rewrite.LegacyIntegration;
using VisioAddIn.Snapping;

namespace ALPS_Visio_AddIn_rewrite
{
    public partial class ThisAddIn
    {
        private LegacyAddInController legacyController;

        /// <summary>
        /// Entrypoint of this Add-In.
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            legacyController = new LegacyAddInController(this, Application);
            legacyController.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            legacyController?.Dispose();
            legacyController = null;
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        internal void updateClicked()
        {
            legacyController.Update();
        }

        internal ModelController getModelController()
        {
            return legacyController.ModelController;
        }

        internal void extendsChanged(SIDPage extends, SIDPage changedPage)
        {
            legacyController.ExtendsChanged(extends, changedPage);
        }

        internal void showDirectoryClicked()
        {
            legacyController.ShowDirectory();
        }

        public void refreshLayerExplorerTreeView()
        {
            legacyController?.RefreshLayerExplorer();
        }
    }
}
