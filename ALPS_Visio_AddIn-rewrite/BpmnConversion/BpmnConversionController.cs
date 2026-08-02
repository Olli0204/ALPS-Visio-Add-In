using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using ALPS_Visio_AddIn_rewrite.Importing;
using PassBpmnConverter.Bpmn;
using PassBpmnConverter.Conversion;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

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

        internal static void Convert(
            string inputFilePath, string outputFilePath)
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
            ValidateSerializedBpmn(outputFilePath);
        }

        internal static void ValidateSerializedBpmn(
            string outputFilePath)
        {
            XDocument document = XDocument.Load(outputFilePath);
            XElement? root = document.Root;
            XNamespace bpmnNamespace =
                BpmnModelConstants.BpmnNs;

            if (root == null
                || root.Name != bpmnNamespace + "definitions")
            {
                throw new InvalidDataException(
                    "Der Export enthält kein gültiges BPMN-definitions-Element.");
            }

            HashSet<string> ids =
                new HashSet<string>(StringComparer.Ordinal);
            foreach (XAttribute idAttribute in root
                .DescendantsAndSelf()
                .Attributes("id"))
            {
                string id = idAttribute.Value;
                try
                {
                    XmlConvert.VerifyNCName(id);
                }
                catch (XmlException exception)
                {
                    throw new InvalidDataException(
                        $"Der BPMN-Export enthält die ungültige ID \"{id}\".",
                        exception);
                }

                if (!ids.Add(id))
                {
                    throw new InvalidDataException(
                        $"Der BPMN-Export enthält die ID \"{id}\" mehrfach.");
                }
            }

            HashSet<string> referenceAttributeNames =
                new HashSet<string>(
                    new[]
                    {
                        "attachedToRef",
                        "bpmnElement",
                        "default",
                        "escalationRef",
                        "messageRef",
                        "processRef",
                        "signalRef",
                        "sourceElement",
                        "sourceRef",
                        "targetElement",
                        "targetRef"
                    },
                    StringComparer.Ordinal);
            IEnumerable<string> references = document
                .Descendants()
                .Attributes()
                .Where(attribute =>
                    referenceAttributeNames.Contains(
                        attribute.Name.LocalName))
                .Select(attribute => attribute.Value)
                .Concat(document
                    .Descendants()
                    .Where(element =>
                        element.Name.LocalName == "incoming"
                        || element.Name.LocalName == "outgoing")
                    .Select(element => element.Value));

            foreach (string reference in references)
            {
                if (!reference.Contains(":")
                    && !ids.Contains(reference))
                {
                    throw new InvalidDataException(
                        $"Der BPMN-Export verweist auf die unbekannte "
                        + $"ID \"{reference}\".");
                }
            }

            if (document
                .Descendants(
                    bpmnNamespace + "conditionalExpression")
                .Any())
            {
                throw new InvalidDataException(
                    "Der BPMN-Export verwendet den ungültigen "
                    + "Elementnamen \"conditionalExpression\".");
            }

            IDictionary<string, XElement> elementsById = root
                .DescendantsAndSelf()
                .Where(element =>
                    element.Attribute("id") != null)
                .ToDictionary(
                    element => element.Attribute("id")!.Value,
                    StringComparer.Ordinal);

            foreach (XElement messageFlow in document
                .Descendants(bpmnNamespace + "messageFlow"))
            {
                string? messageFlowName =
                    messageFlow.Attribute("name")?.Value;
                string? messageReference =
                    messageFlow.Attribute("messageRef")?.Value;
                if (string.IsNullOrWhiteSpace(messageFlowName)
                    || string.IsNullOrWhiteSpace(messageReference)
                    || !elementsById.TryGetValue(
                        messageReference,
                        out XElement message))
                {
                    continue;
                }

                string? messageName =
                    message.Attribute("name")?.Value;
                if (string.Equals(
                        messageFlowName,
                        messageName,
                        StringComparison.Ordinal))
                {
                    throw new InvalidDataException(
                        "Der BPMN-Export beschriftet einen "
                        + "Nachrichtenfluss und seine Nachricht doppelt "
                        + $"mit \"{messageFlowName}\".");
                }
            }

            foreach (XElement gateway in document
                .Descendants(
                    bpmnNamespace + "exclusiveGateway"))
            {
                IList<string> outgoingIds = gateway
                    .Elements(bpmnNamespace + "outgoing")
                    .Select(element => element.Value)
                    .ToList();
                if (outgoingIds.Count < 2)
                    continue;

                string? defaultId =
                    gateway.Attribute("default")?.Value;
                foreach (string outgoingId in outgoingIds)
                {
                    if (outgoingId == defaultId)
                        continue;

                    if (!elementsById.TryGetValue(
                            outgoingId,
                            out XElement sequenceFlow)
                        || sequenceFlow.Element(
                            bpmnNamespace
                            + "conditionExpression") == null)
                    {
                        throw new InvalidDataException(
                            $"Der BPMN-Export enthält für die "
                            + $"Abzweigung \"{outgoingId}\" keine "
                            + "Bedingung.");
                    }
                }
            }

            foreach (XAttribute coordinate in document
                .Descendants()
                .Attributes()
                .Where(attribute =>
                    attribute.Name.LocalName == "x"
                    || attribute.Name.LocalName == "y"
                    || attribute.Name.LocalName == "width"
                    || attribute.Name.LocalName == "height"))
            {
                if (!double.TryParse(
                        coordinate.Value,
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out double value)
                    || double.IsNaN(value)
                    || double.IsInfinity(value))
                {
                    throw new InvalidDataException(
                        $"Der BPMN-Export enthält die ungültige "
                        + $"Diagrammkoordinate \"{coordinate.Value}\".");
                }
            }

            XNamespace dcNamespace =
                BpmnModelConstants.OmgDcNs;
            XNamespace diNamespace =
                BpmnModelConstants.OmgDiNs;
            XNamespace bpmnDiNamespace =
                BpmnModelConstants.BpmnDiNs;
            double maximumShapeExtent = document
                .Descendants(dcNamespace + "Bounds")
                .SelectMany(bounds =>
                {
                    double x = ReadCoordinate(bounds, "x");
                    double y = ReadCoordinate(bounds, "y");
                    double width = ReadCoordinate(bounds, "width");
                    double height = ReadCoordinate(bounds, "height");
                    return new[]
                    {
                        Math.Abs(x),
                        Math.Abs(y),
                        Math.Abs(x + width),
                        Math.Abs(y + height)
                    };
                })
                .DefaultIfEmpty(0)
                .Max();
            double maximumWaypointCoordinate =
                Math.Max(10000, maximumShapeExtent * 10);

            foreach (XElement waypoint in document
                .Descendants(diNamespace + "waypoint"))
            {
                double x = ReadCoordinate(waypoint, "x");
                double y = ReadCoordinate(waypoint, "y");
                if (Math.Abs(x) > maximumWaypointCoordinate
                    || Math.Abs(y) > maximumWaypointCoordinate)
                {
                    throw new InvalidDataException(
                        "Der BPMN-Export enthält eine Diagrammkante "
                        + "außerhalb des sichtbaren Modellbereichs.");
                }
            }

            foreach (XElement edge in document
                .Descendants(bpmnDiNamespace + "BPMNEdge"))
            {
                IList<XElement> waypoints = edge
                    .Elements(diNamespace + "waypoint")
                    .ToList();
                if (waypoints.Count < 2)
                {
                    throw new InvalidDataException(
                        "Der BPMN-Export enthält eine Diagrammkante "
                        + "mit weniger als zwei Wegpunkten.");
                }

                for (int index = 1; index < waypoints.Count; index++)
                {
                    double previousX = ReadCoordinate(
                        waypoints[index - 1],
                        "x");
                    double previousY = ReadCoordinate(
                        waypoints[index - 1],
                        "y");
                    double currentX = ReadCoordinate(
                        waypoints[index],
                        "x");
                    double currentY = ReadCoordinate(
                        waypoints[index],
                        "y");
                    bool horizontal =
                        Math.Abs(previousY - currentY) < 0.001;
                    bool vertical =
                        Math.Abs(previousX - currentX) < 0.001;
                    if (!horizontal && !vertical)
                    {
                        throw new InvalidDataException(
                            "Der BPMN-Export enthält eine diagonale "
                            + "Diagrammkante statt eines "
                            + "orthogonalen Routings.");
                    }
                }
            }
        }

        private static double ReadCoordinate(
            XElement element,
            string attributeName)
        {
            XAttribute? attribute = element.Attribute(attributeName);
            return attribute == null
                ? 0
                : double.Parse(
                    attribute.Value,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture);
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
