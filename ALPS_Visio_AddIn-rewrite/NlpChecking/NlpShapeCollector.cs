using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    /// <summary>
    /// Copies the supported ShapeSheet values while execution is still on the
    /// Visio UI thread. No COM object escapes this collector.
    /// </summary>
    internal static class NlpShapeCollector
    {
        private static readonly string[] SupportedTypes =
        {
            "FullySpecifiedSubject",
            "MultiSubject",
            "DoState",
            "MessageSpecification",
            "DoTransition",
            "SendState",
            "ReceiveState",
            "InterfaceSubject"
        };

        public static IList<NlpShapeCandidate> Collect(Visio.Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            List<NlpShapeCandidate> candidates =
                new List<NlpShapeCandidate>();
            foreach (Visio.Page page in document.Pages)
            {
                foreach (Visio.Shape shape in page.Shapes)
                {
                    string shapeType = ResolveShapeType(shape);
                    if (shapeType == null)
                        continue;

                    string label = ReadCell(
                        shape, "Prop." + Constants.Properties.Label);
                    if (string.IsNullOrWhiteSpace(label))
                        label = ReadShapeText(shape);

                    candidates.Add(new NlpShapeCandidate
                    {
                        PageName = page.Name,
                        ShapeId = shape.ID,
                        ShapeName = shape.Name,
                        Label = label ?? string.Empty,
                        ShapeType = shapeType
                    });
                }
            }

            return candidates;
        }

        private static string ResolveShapeType(Visio.Shape shape)
        {
            if (IsTrue(ReadCell(shape,
                "Prop." + Constants.Properties.Subject.Multi)))
            {
                return "MultiSubject";
            }

            string rawType = ReadCell(shape, "Prop.modelComponentType");
            string masterName = ReadMasterName(shape);
            string combined = (rawType ?? string.Empty)
                + " " + (masterName ?? string.Empty);

            if (combined.IndexOf(
                Constants.SIDMasters.StandardActor,
                StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "FullySpecifiedSubject";
            }

            if (combined.IndexOf(
                Constants.SIDMasters.InterfaceActor,
                StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "InterfaceSubject";
            }

            if (combined.IndexOf(
                Constants.SBDMasters.DoState,
                StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "DoState";
            }

            combined = combined.Replace("Recieve", "Receive");
            foreach (string supportedType in SupportedTypes)
            {
                if (combined.IndexOf(
                    supportedType, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return supportedType;
                }
            }

            return null;
        }

        private static string ReadCell(Visio.Shape shape, string cellName)
        {
            try
            {
                if (shape.CellExistsU[cellName, 0] == 0)
                    return null;

                return shape.CellsU[cellName]
                    .ResultStr[(short)Visio.VisUnitCodes.visNoCast]
                    ?.Trim().Trim('"');
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static string ReadShapeText(Visio.Shape shape)
        {
            try
            {
                return string.IsNullOrWhiteSpace(shape.Text)
                    ? shape.Name
                    : shape.Text;
            }
            catch (COMException)
            {
                return string.Empty;
            }
        }

        private static string ReadMasterName(Visio.Shape shape)
        {
            try
            {
                return shape.Master?.NameU;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static bool IsTrue(string value)
        {
            return string.Equals(value, "TRUE",
                       StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "-1",
                       StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "1",
                       StringComparison.OrdinalIgnoreCase);
        }
    }
}
