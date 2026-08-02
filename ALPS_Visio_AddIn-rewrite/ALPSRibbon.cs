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

            RibbonSplitButton verificationButton =
                CreateVerificationButton();
            owlGroup.Items.Add(verificationButton);

            RibbonGroup conversionGroup = this.Factory.CreateRibbonGroup();
            conversionGroup.Name = "modelConversionGroup";
            conversionGroup.Label = "Model Conversion";
            alpsTab.Groups.Add(conversionGroup);

            RibbonSplitButton passToBpmnButton =
                CreateBpmnConversionButton();
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
            autoArrangeButton.Image = Properties.Resources.autoArrange;
            autoArrangeButton.ShowLabel = true;
            autoArrangeButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            autoArrangeButton.ItemSize = RibbonControlSize.RibbonControlSizeRegular;
            autoArrangeButton.Click += new RibbonControlEventHandler(this.AutoArrangeTopDown);

            RibbonButton topDownButton = this.Factory.CreateRibbonButton();
            topDownButton.Name = "autoArrangeTopDownButton";
            topDownButton.Label = "Top-down";
            topDownButton.ScreenTip = "Von oben nach unten anordnen";
            topDownButton.Image = Properties.Resources.autoArrangeTopDown;
            topDownButton.ShowImage = true;
            topDownButton.Click += new RibbonControlEventHandler(this.AutoArrangeTopDown);
            autoArrangeButton.Items.Add(topDownButton);

            RibbonButton leftRightButton = this.Factory.CreateRibbonButton();
            leftRightButton.Name = "autoArrangeLeftRightButton";
            leftRightButton.Label = "Left-right";
            leftRightButton.ScreenTip = "Von links nach rechts anordnen";
            leftRightButton.Image = Properties.Resources.autoArrangeLeftRight;
            leftRightButton.ShowImage = true;
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
            apiSettingsButton.Image =
                Properties.Resources.nlpProviderSettings;
            apiSettingsButton.ShowImage = true;
            apiSettingsButton.Click += new RibbonControlEventHandler(
                this.ConfigureNlpApi);
            button.Items.Add(apiSettingsButton);

            return button;
        }

        private RibbonSplitButton CreateVerificationButton()
        {
            RibbonSplitButton button =
                this.Factory.CreateRibbonSplitButton();
            button.Name = "verificationButton";
            button.Label = "Verify ALPS Models";
            button.ScreenTip = "Verify an ALPS implementation";
            button.SuperTip = "Compare an abstract ALPS specification "
                + "with an implementation. Use the menu to export and "
                + "verify the currently open Visio model directly.";
            button.OfficeImageId = "AdpDiagramArrangeTables";
            button.ShowLabel = true;
            button.ControlSize =
                RibbonControlSize.RibbonControlSizeLarge;
            button.ItemSize =
                RibbonControlSize.RibbonControlSizeRegular;
            button.Click += new RibbonControlEventHandler(
                this.RunAlpsVerification);

            RibbonButton filesButton =
                this.Factory.CreateRibbonButton();
            filesButton.Name = "verificationFilesButton";
            filesButton.Label = "OWL/RDF-Dateien auswählen";
            filesButton.ScreenTip =
                "Spezifikation und Implementierung aus Dateien prüfen";
            filesButton.Click += new RibbonControlEventHandler(
                this.RunAlpsVerification);
            button.Items.Add(filesButton);

            RibbonButton currentImplementationButton =
                this.Factory.CreateRibbonButton();
            currentImplementationButton.Name =
                "verificationCurrentImplementationButton";
            currentImplementationButton.Label =
                "Aktuelles Modell als Implementierung";
            currentImplementationButton.ScreenTip =
                "Aktuelle Visio-Zeichnung per OWL-Makro exportieren";
            currentImplementationButton.Click +=
                new RibbonControlEventHandler(
                    this.RunAlpsVerificationWithCurrentImplementation);
            button.Items.Add(currentImplementationButton);

            RibbonButton currentSpecificationButton =
                this.Factory.CreateRibbonButton();
            currentSpecificationButton.Name =
                "verificationCurrentSpecificationButton";
            currentSpecificationButton.Label =
                "Aktuelles Modell als Spezifikation";
            currentSpecificationButton.ScreenTip =
                "Aktuelle Visio-Zeichnung per OWL-Makro exportieren";
            currentSpecificationButton.Click +=
                new RibbonControlEventHandler(
                    this.RunAlpsVerificationWithCurrentSpecification);
            button.Items.Add(currentSpecificationButton);

            return button;
        }

        private RibbonSplitButton CreateBpmnConversionButton()
        {
            RibbonSplitButton button =
                this.Factory.CreateRibbonSplitButton();
            button.Name = "passToBpmnButton";
            button.Label = "Convert PASS to BPMN";
            button.ScreenTip = "Convert a PASS model to BPMN";
            button.SuperTip = "Convert a selected PASS/ALPS OWL file or "
                + "export the currently open Visio model through the SID "
                + "stencil's OWL exporter before converting it.";
            button.OfficeImageId = "FileSaveAs";
            button.ShowLabel = true;
            button.ControlSize =
                RibbonControlSize.RibbonControlSizeLarge;
            button.ItemSize =
                RibbonControlSize.RibbonControlSizeRegular;
            button.Click += new RibbonControlEventHandler(
                this.ConvertPassToBpmn);

            RibbonButton fileButton =
                this.Factory.CreateRibbonButton();
            fileButton.Name = "passToBpmnFileButton";
            fileButton.Label = "OWL/RDF-Datei auswählen";
            fileButton.Click += new RibbonControlEventHandler(
                this.ConvertPassToBpmn);
            button.Items.Add(fileButton);

            RibbonButton currentModelButton =
                this.Factory.CreateRibbonButton();
            currentModelButton.Name = "passToBpmnCurrentModelButton";
            currentModelButton.Label = "Aktuelles Visio-Modell";
            currentModelButton.ScreenTip =
                "Aktuelle Zeichnung per OWL-Makro exportieren und "
                + "nach BPMN konvertieren";
            currentModelButton.Click += new RibbonControlEventHandler(
                this.ConvertCurrentVisioModelToBpmn);
            button.Items.Add(currentModelButton);

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

        private void RunAlpsVerificationWithCurrentImplementation(
            object sender, RibbonControlEventArgs e)
        {
            verificationController
                .RunWithCurrentModelAsImplementation();
        }

        private void RunAlpsVerificationWithCurrentSpecification(
            object sender, RibbonControlEventArgs e)
        {
            verificationController
                .RunWithCurrentModelAsSpecification();
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

        private void ConvertCurrentVisioModelToBpmn(
            object sender, RibbonControlEventArgs e)
        {
            bpmnConversionController.RunFromCurrentVisioModel();
        }
    }
}
