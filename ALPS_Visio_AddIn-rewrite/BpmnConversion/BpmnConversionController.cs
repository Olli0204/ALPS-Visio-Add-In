using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using ALPS_Visio_AddIn_rewrite.Importing;
using PassBpmnConverter.Bpmn;
using PassBpmnConverter.Conversion;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ALPS_Visio_AddIn_rewrite.BpmnConversion
{
    /// <summary>
    /// Coordinates PASS OWL selection, conversion, and BPMN file output.
    /// </summary>
    internal sealed class BpmnConversionController
    {
        public void Run()
        {
            string? inputFilePath = SelectInputFile();
            if (inputFilePath == null) return;

            string? outputFilePath = SelectOutputFile(inputFilePath);
            if (outputFilePath == null) return;

            Cursor previousCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                Convert(inputFilePath, outputFilePath);

                MessageBox.Show(
                    "Das PASS-Modell wurde erfolgreich nach BPMN konvertiert:\n\n"
                    + outputFilePath,
                    "PASS nach BPMN",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    "Das PASS-Modell konnte nicht nach BPMN konvertiert werden.\n\n"
                    + exception.Message,
                    "PASS nach BPMN",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = previousCursor;
            }
        }

        private static void Convert(string inputFilePath, string outputFilePath)
        {
            IPASSReaderWriter parser =
                OwlImportComposition.CreateConfiguredParser();
            IList<IPASSProcessModel> models = parser.loadModels(
                new List<string> { inputFilePath });
            IPASSProcessModel? passModel =
                models == null ? null : models.FirstOrDefault();

            if (passModel == null)
                throw new InvalidDataException(
                    "Die ausgewählte Datei enthält kein lesbares PASS-Modell.");

            IBpmnModel bpmnModel =
                Converter.ConvertPassToBpmn(passModel);
            BpmnDiagramGenerator.GenerateDiagram(bpmnModel);
            BpmnSerializer.Serialize(bpmnModel, outputFilePath);
        }

        private static string? SelectInputFile()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "PASS-Modell für die BPMN-Konvertierung auswählen";
                dialog.Filter =
                    "PASS Ontology Files (*.owl;*.rdf)|*.owl;*.rdf"
                    + "|OWL Files (*.owl)|*.owl"
                    + "|RDF Files (*.rdf)|*.rdf"
                    + "|All Files (*.*)|*.*";

                return dialog.ShowDialog() == DialogResult.OK
                    ? dialog.FileName
                    : null;
            }
        }

        private static string? SelectOutputFile(string inputFilePath)
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "BPMN-Modell speichern";
                dialog.Filter = "BPMN 2.0 Files (*.bpmn)|*.bpmn";
                dialog.DefaultExt = "bpmn";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;
                dialog.InitialDirectory =
                    Path.GetDirectoryName(inputFilePath);
                dialog.FileName =
                    Path.GetFileNameWithoutExtension(inputFilePath) + ".bpmn";

                return dialog.ShowDialog() == DialogResult.OK
                    ? dialog.FileName
                    : null;
            }
        }
    }
}
