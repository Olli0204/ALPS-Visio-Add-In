using ALPS_Visio_AddIn_rewrite.BpmnConversion;
using ALPS_Visio_AddIn_rewrite.NlpChecking;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Core;
using System.Windows.Forms;
using ALPS_Visio_AddIn_rewrite.Verification;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite
{
    partial class ALPSRibbon : RibbonBase
    {
        private readonly OWLImporter owlImporter;
        private readonly BpmnConversionController bpmnConversionController;
        private readonly NlpCheckingController nlpCheckingController;
        private readonly AlpsVerificationController
            verificationController;

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
            this.bpmnConversionController =
                new BpmnConversionController();
            this.nlpCheckingController = new NlpCheckingController(
                () => Globals.ThisAddIn.GetDrawingDocument());
            this.verificationController =
                new AlpsVerificationController();
            this.RibbonType = "Microsoft.Visio.Drawing";

            RibbonTab alpsTab = this.Factory.CreateRibbonTab();
            alpsTab.Name = "alpsPassTab";
            alpsTab.Label = "ALPS/PASS ADDIN";
            this.Tabs.Add(alpsTab);

            RibbonGroup standardGroup = this.Factory.CreateRibbonGroup();
            standardGroup.Name = "standardFunctionsGroup";
            standardGroup.Label = "Standard Functions";
            alpsTab.Groups.Add(standardGroup);

            RibbonButton openStencilsButton = this.Factory.CreateRibbonButton();
            openStencilsButton.Name = "openStencilsButton";
            openStencilsButton.Label = "Open ALPS/PASS Stencils";
            openStencilsButton.SuperTip = "Tries to open the necessary ALPS Visio stencils if they are available on the system.";
            openStencilsButton.Image = Properties.Resources.document_open_7;
            openStencilsButton.ShowImage = true;
            openStencilsButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            openStencilsButton.Click += new RibbonControlEventHandler(this.OpenStencils);
            standardGroup.Items.Add(openStencilsButton);

            RibbonGroup layerEditingGroup = this.Factory.CreateRibbonGroup();
            layerEditingGroup.Name = "layerEditingGroup";
            layerEditingGroup.Label = "ALPS Layer Editing";
            alpsTab.Groups.Add(layerEditingGroup);

            RibbonButton layerExplorerButton = this.Factory.CreateRibbonButton();
            layerExplorerButton.Name = "layerExplorerButton";
            layerExplorerButton.Label = "Show layer Explorer";
            layerExplorerButton.SuperTip = "Open the layer explorer for advanced multi-layered ALPS editing.";
            layerExplorerButton.OfficeImageId = "LayersMenu";
            layerExplorerButton.ShowImage = true;
            layerExplorerButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            layerExplorerButton.Click += new RibbonControlEventHandler(this.ShowLayerExplorer);
            layerEditingGroup.Items.Add(layerExplorerButton);

            RibbonGroup owlGroup = this.Factory.CreateRibbonGroup();
            owlGroup.Name = "owlPassToolsGroup";
            owlGroup.Label = "OWL PASS Tools";
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

            RibbonSplitButton autoArrangeButton = CreateAutoArrangeButton();
            owlGroup.Items.Add(autoArrangeButton);

            RibbonButton verificationButton = this.Factory.CreateRibbonButton();
            verificationButton.Name = "verificationButton";
            verificationButton.Label = "Verify ALPS Models";
            verificationButton.SuperTip = "Select an abstract ALPS specification and an implementing OWL/RDF model. The add-in checks the SID implementation relationships, fully specified subjects, and communication restrictions supported by the ALPS verification thesis prototype.";
            verificationButton.OfficeImageId = "AdpDiagramArrangeTables";
            verificationButton.ShowImage = true;
            verificationButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            verificationButton.Click += new RibbonControlEventHandler(
                this.RunAlpsVerification);
            owlGroup.Items.Add(verificationButton);

            RibbonGroup conversionGroup = this.Factory.CreateRibbonGroup();
            conversionGroup.Name = "modelConversionGroup";
            conversionGroup.Label = "Model Conversion";
            alpsTab.Groups.Add(conversionGroup);

            RibbonButton passToBpmnButton =
                this.Factory.CreateRibbonButton();
            passToBpmnButton.Name = "passToBpmnButton";
            passToBpmnButton.Label = "Convert PASS to BPMN";
            passToBpmnButton.ScreenTip = "Convert a PASS OWL model to BPMN";
            passToBpmnButton.SuperTip =
                "Select a PASS or ALPS OWL file and save the converted "
                + "model as a BPMN 2.0 file.";
            passToBpmnButton.OfficeImageId = "FileSaveAs";
            passToBpmnButton.ShowImage = true;
            passToBpmnButton.ControlSize =
                RibbonControlSize.RibbonControlSizeLarge;
            passToBpmnButton.Click += new RibbonControlEventHandler(
                this.ConvertPassToBpmn);
            conversionGroup.Items.Add(passToBpmnButton);

            RibbonGroup nlpGroup = this.Factory.CreateRibbonGroup();
            nlpGroup.Name = "nlpPassCheckingGroup";
            nlpGroup.Label = "NLP PASS Checking";
            alpsTab.Groups.Add(nlpGroup);

            RibbonSplitButton nlpCheckingButton =
                CreateNlpCheckingButton();
            nlpGroup.Items.Add(nlpCheckingButton);
        }

        private RibbonSplitButton CreateAutoArrangeButton()
        {
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

            return autoArrangeButton;
        }

        private RibbonSplitButton CreateNlpCheckingButton()
        {
            RibbonSplitButton button =
                this.Factory.CreateRibbonSplitButton();
            button.Name = "nlpCheckingButton";
            button.Label = "Check Model Naming";
            button.ScreenTip = "Check PASS model labels";
            button.SuperTip = "Checks supported PASS shape labels with "
                + "the locally trained classifier. Optional naming "
                + "suggestions can be enabled in Provider Settings.";
            button.OfficeImageId = "Spelling";
            button.ShowLabel = true;
            button.ControlSize =
                RibbonControlSize.RibbonControlSizeLarge;
            button.ItemSize =
                RibbonControlSize.RibbonControlSizeRegular;
            button.Click += new RibbonControlEventHandler(
                this.CheckModelNaming);

            RibbonButton checkButton =
                this.Factory.CreateRibbonButton();
            checkButton.Name = "nlpCheckModelNamingButton";
            checkButton.Label = "Check Model Naming";
            checkButton.Click += new RibbonControlEventHandler(
                this.CheckModelNaming);
            button.Items.Add(checkButton);

            RibbonButton retrainButton =
                this.Factory.CreateRibbonButton();
            retrainButton.Name = "nlpRetrainButton";
            retrainButton.Label = "Retrain";
            retrainButton.ScreenTip =
                "Retrain from the bundled PASS examples";
            retrainButton.Click += new RibbonControlEventHandler(
                this.RetrainNlpClassifier);
            button.Items.Add(retrainButton);

            RibbonButton apiSettingsButton =
                this.Factory.CreateRibbonButton();
            apiSettingsButton.Name = "nlpApiSettingsButton";
            apiSettingsButton.Label = "Provider Settings";
            apiSettingsButton.ScreenTip =
                "Configure providers, API keys, and models";
            apiSettingsButton.Click += new RibbonControlEventHandler(
                this.ConfigureNlpApi);
            button.Items.Add(apiSettingsButton);

            return button;
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
            VisioHelper.OpenInteractiveStencils();
        }

        /// <summary>
        /// Show the layer explorer.
        /// </summary>
        private void ShowLayerExplorer(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.showDirectoryClicked();
        }

        private void RunAlpsVerification(
            object sender, RibbonControlEventArgs e)
        {
            verificationController.Run();
        }

        private async void CheckModelNaming(
            object sender, RibbonControlEventArgs e)
        {
            await nlpCheckingController.RunAsync();
        }

        private void RetrainNlpClassifier(
            object sender, RibbonControlEventArgs e)
        {
            nlpCheckingController.Retrain();
        }

        private void ConfigureNlpApi(
            object sender, RibbonControlEventArgs e)
        {
            nlpCheckingController.ConfigureProviders();
        }

        private void ConvertPassToBpmn(
            object sender, RibbonControlEventArgs e)
        {
            bpmnConversionController.Run();
        }
    }
}
