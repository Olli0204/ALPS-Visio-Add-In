using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Core;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite
{
    partial class ALPSRibbon : RibbonBase
    {
        private readonly OWLImporter owlImporter;

        /// <summary>
        /// Ribbon containing ALPS Menu
        /// </summary>
        public ALPSRibbon()
            : this(new OWLImporter())
        {
        }

        internal ALPSRibbon(OWLImporter owlImporter)
            : base(Globals.Factory.GetRibbonFactory())
        {
            this.owlImporter = owlImporter
                ?? throw new System.ArgumentNullException(nameof(owlImporter));
            this.RibbonType = "Microsoft.Visio.Drawing";

            RibbonTab alpsTab = this.Factory.CreateRibbonTab();
            alpsTab.Label = "ALPS/PASS ADDIN";
            this.Tabs.Add(alpsTab);

            RibbonGroup owlGroup = this.Factory.CreateRibbonGroup();
            owlGroup.Label = "ALPS Tools";
            alpsTab.Groups.Add(owlGroup);

            RibbonButton owlImporterButton = this.Factory.CreateRibbonButton();
            owlImporterButton.Name = "owlImporterButton";
            owlImporterButton.Label = "Import OWL";
            owlImporterButton.SuperTip = "Use this tool to import PASS and ALPS Process Models from OWL Files based on the standard pass ontology";
            owlImporterButton.Image = Properties.Resources.owlIcon2;
            owlImporterButton.ShowImage = true;
            owlImporterButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            owlImporterButton.Click += new RibbonControlEventHandler(this.LoadOWLFile);
            owlGroup.Items.Add(owlImporterButton);

            RibbonSplitButton autoArrangeButton = this.Factory.CreateRibbonSplitButton();
            autoArrangeButton.Name = "autoArrangeButton";
            autoArrangeButton.Label = "Auto-Arrange";
            autoArrangeButton.ScreenTip = "Graph automatisch anordnen";
            autoArrangeButton.SuperTip = "Ordnet den Graphen standardmäßig von oben nach unten an. Über das Menü kann alternativ eine Anordnung von links nach rechts gewählt werden.";
            autoArrangeButton.Image = Properties.Resources.pageSetup;
            autoArrangeButton.ShowLabel = true;
            autoArrangeButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            autoArrangeButton.ItemSize = RibbonControlSize.RibbonControlSizeRegular;
            autoArrangeButton.Click += new RibbonControlEventHandler(this.AutoArrangeTopDown);

            RibbonButton topDownButton = this.Factory.CreateRibbonButton();
            topDownButton.Name = "autoArrangeTopDownButton";
            topDownButton.Label = "Top-down";
            topDownButton.ScreenTip = "Von oben nach unten anordnen";
            topDownButton.Click += new RibbonControlEventHandler(this.AutoArrangeTopDown);
            autoArrangeButton.Items.Add(topDownButton);

            RibbonButton leftRightButton = this.Factory.CreateRibbonButton();
            leftRightButton.Name = "autoArrangeLeftRightButton";
            leftRightButton.Label = "Left-right";
            leftRightButton.ScreenTip = "Von links nach rechts anordnen";
            leftRightButton.Click += new RibbonControlEventHandler(this.AutoArrangeLeftRight);
            autoArrangeButton.Items.Add(leftRightButton);

            owlGroup.Items.Add(autoArrangeButton);

            RibbonButton openStencilsButton = this.Factory.CreateRibbonButton();
            openStencilsButton.Name = "openStencilsButton";
            openStencilsButton.Label = "Open ALPS/PASS Stencils";
            openStencilsButton.SuperTip = "Tries to open the (necessary) ALPS Visio stencils if they are available on the system.";
            openStencilsButton.Image = Properties.Resources.document_open_7;
            openStencilsButton.ShowImage = true;
            openStencilsButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            openStencilsButton.Click += new RibbonControlEventHandler(this.OpenStencils);
            owlGroup.Items.Add(openStencilsButton);

            RibbonButton layerExplorerButton = this.Factory.CreateRibbonButton();
            layerExplorerButton.Name = "layerExplorerButton";
            layerExplorerButton.Label = "Show layer Explorer";
            layerExplorerButton.SuperTip = "Open a the layer explorer, a tool for advanced multi-layered ALPS (Abstract Layered PASS editing)";
            layerExplorerButton.Image = Properties.Resources.pageSetup;
            layerExplorerButton.ShowImage = true;
            layerExplorerButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            layerExplorerButton.Click += new RibbonControlEventHandler(this.ShowLayerExplorer);
            owlGroup.Items.Add(layerExplorerButton);

            // FEAT: ALPS verification tool
            // FEAT: PASS natural language checker
            // FEAT: PASS BPMN converter
        }

        /// <summary>
        /// Open file dialog and import OWL file.
        /// </summary>
        private void LoadOWLFile(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Ontology Files (.owl)|*.owl|RDF Files (*.rdf)|*.rdf"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
                owlImporter.Parse(dialog.FileName);
        }

        private void AutoArrangeTopDown(object sender, RibbonControlEventArgs e)
        {
            AutoArrange(VisioHelper.GraphLayoutDirection.TopDown);
        }

        private void AutoArrangeLeftRight(object sender, RibbonControlEventArgs e)
        {
            AutoArrange(VisioHelper.GraphLayoutDirection.LeftRight);
        }

        private static void AutoArrange(VisioHelper.GraphLayoutDirection direction)
        {
            Visio.IVPage activePage = Globals.ThisAddIn.Application.ActivePage;
            if (activePage == null)
            {
                MessageBox.Show("Es ist keine Zeichnungsseite aktiv.", "ALPS/PASS Auto-Arrange",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                VisioHelper.AutoArrangePage(activePage, direction);
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                MessageBox.Show("Der Graph konnte nicht automatisch angeordnet werden.\n\n" + exception.Message,
                    "ALPS/PASS Auto-Arrange", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Open ALPS stencils.
        /// </summary>
        private void OpenStencils(object sender, RibbonControlEventArgs e)
        {
            VisioHelper.openStencil(VisioHelper.VisioStencils.SID_STENCIL);
        }

        /// <summary>
        /// Show the layer explorer.
        /// </summary>
        private void ShowLayerExplorer(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.showDirectoryClicked();
        }
    }
}
