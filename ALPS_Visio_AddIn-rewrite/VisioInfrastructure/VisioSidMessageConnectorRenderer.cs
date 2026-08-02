using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Visio = Microsoft.Office.Interop.Visio;
using LayoutDirection =
    ALPS_Visio_AddIn_rewrite.VisioHelper.GraphLayoutDirection;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Separates the semantic SID message master from its visible route.
    /// </summary>
    internal static class VisioSidMessageConnectorRenderer
    {
        private static readonly string[] LineCells =
        {
            "LineWeight",
            "LineColor",
            "LineColorTrans",
            "LinePattern",
            "Rounding",
            "EndArrowSize",
            "EndArrow",
            "CompoundType"
        };

        public static void Ensure(
            Visio.IVPage page, LayoutDirection direction)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));

            List<Visio.Shape> semanticConnectors =
                new List<Visio.Shape>();
            foreach (Visio.Shape shape in page.Shapes)
            {
                // Include an existing shadow as well. If its generated visual
                // connector was removed manually, the next Auto-Arrange run
                // must be able to recreate it from the semantic master.
                if (IsVisualConnector(shape)
                    || !VisioConnectorRebinder.IsStencilSidMessageConnector(
                        shape))
                {
                    continue;
                }

                semanticConnectors.Add(shape);
            }

            foreach (Visio.Shape semanticConnector in semanticConnectors)
            {
                if (!VisioConnectorRebinder.TryGetConnectedShapes(
                    page, semanticConnector,
                    out Visio.Shape source, out Visio.Shape target))
                {
                    continue;
                }

                Visio.Shape visualConnector = FindVisualConnector(
                    page, semanticConnector.ID);
                bool created = visualConnector == null;
                if (visualConnector == null)
                {
                    visualConnector = CreateVisualConnector(
                        page, semanticConnector);
                }

                if (visualConnector == null) continue;

                MarkVisualConnector(
                    visualConnector, semanticConnector, source, target);
                if (!VisioConnectorRebinder.RebindKnownConnector(
                    page, visualConnector, source, target, direction))
                {
                    if (created) TryDelete(visualConnector);
                    continue;
                }

                HideSemanticConnector(page, semanticConnector);
                try
                {
                    visualConnector.SendToBack();
                }
                catch (COMException)
                {
                    // Z-order is cosmetic. The connector remains functional
                    // in protected documents that reject the operation.
                }
            }
        }

        internal static bool IsVisualConnector(Visio.Shape shape)
        {
            return GetUserFlag(
                shape, Constants.UserCells.AutoArrangeSidVisualConnector);
        }

        internal static bool IsSemanticShadow(Visio.Shape shape)
        {
            return GetUserFlag(
                shape, Constants.UserCells.AutoArrangeSidSemanticShadow);
        }

        private static Visio.Shape FindVisualConnector(
            Visio.IVPage page, int semanticConnectorId)
        {
            string expectedId = semanticConnectorId.ToString(
                CultureInfo.InvariantCulture);
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (IsVisualConnector(shape)
                    && string.Equals(GetUserValue(shape,
                        Constants.UserCells
                            .AutoArrangeSidSemanticConnectorShapeId),
                        expectedId, StringComparison.OrdinalIgnoreCase))
                {
                    return shape;
                }
            }

            return null;
        }

        private static Visio.Shape CreateVisualConnector(
            Visio.IVPage page, Visio.Shape semanticConnector)
        {
            try
            {
                object connectorTool =
                    page.Application.ConnectorToolDataObject;
                Visio.Shape connector = page.Drop(connectorTool, 0, 0);
                CopyLineFormatting(semanticConnector, connector);
                VisioRouting.TrySetCell(
                    connector, "ConFixedCode", 0, false);
                VisioRouting.TrySetCell(
                    connector, "ShapeRouteStyle", 1, false);
                VisioRouting.TrySetCell(
                    connector, "ConLineRouteExt", 1, false);
                VisioRouting.TrySetCell(
                    connector, "ConLineJumpCode", 0, false);
                VisioRouting.TrySetCell(
                    connector, "BeginArrow", 0, false);
                return connector;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static void MarkVisualConnector(Visio.Shape visualConnector,
            Visio.Shape semanticConnector, Visio.Shape source,
            Visio.Shape target)
        {
            VisioShapeSheet.SetUserCell(visualConnector,
                Constants.UserCells.AutoArrangeSidVisualConnector, 1);
            VisioShapeSheet.SetUserCell(visualConnector,
                Constants.UserCells.AutoArrangeSidSemanticConnectorShapeId,
                semanticConnector.ID);
            VisioShapeSheet.SetUserCell(visualConnector,
                Constants.UserCells.AutoArrangeSourceShapeId, source.ID);
            VisioShapeSheet.SetUserCell(visualConnector,
                Constants.UserCells.AutoArrangeTargetShapeId, target.ID);
        }

        private static void HideSemanticConnector(
            Visio.IVPage page, Visio.Shape semanticConnector)
        {
            VisioShapeSheet.SetUserCell(semanticConnector,
                Constants.UserCells.AutoArrangeSidSemanticShadow, 1);

            try
            {
                Visio.Layer layer;
                try
                {
                    layer = page.Layers.ItemU[
                        Constants.Layers.InternalSidSemantics];
                }
                catch (COMException)
                {
                    layer = page.Layers.Add(
                        Constants.Layers.InternalSidSemantics);
                }

                layer.CellsC[4].FormulaU = "0";
                layer.CellsC[5].FormulaU = "0";
                layer.Add(semanticConnector, (short)1);
            }
            catch (COMException)
            {
                // If a protected document rejects layers, make the complete
                // group transparent instead of leaving duplicate geometry.
                HideLineRecursively(semanticConnector);
            }
        }

        private static void HideLineRecursively(Visio.Shape shape)
        {
            TrySetFormula(shape, "LineColorTrans", "100%");
            try
            {
                foreach (Visio.Shape child in shape.Shapes)
                    HideLineRecursively(child);
            }
            catch (COMException)
            {
                // A non-group shape has no child collection to hide.
            }
        }

        private static void CopyLineFormatting(
            Visio.Shape source, Visio.Shape target)
        {
            foreach (string cellName in LineCells)
            {
                try
                {
                    if (source.CellExistsU[cellName, 0] == 0
                        || target.CellExistsU[cellName, 0] == 0)
                    {
                        continue;
                    }

                    double value = source.CellsU[cellName].Result[""];
                    target.CellsU[cellName].FormulaU =
                        value.ToString(CultureInfo.InvariantCulture);
                }
                catch (COMException)
                {
                    // Formatting is best-effort; endpoint routing is not.
                }
            }
        }

        private static bool GetUserFlag(Visio.Shape shape, string rowName)
        {
            string value = GetUserValue(shape, rowName);
            return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "TRUE",
                    StringComparison.OrdinalIgnoreCase);
        }

        private static string GetUserValue(
            Visio.Shape shape, string rowName)
        {
            if (shape == null) return null;
            string cellName = "User." + rowName;
            try
            {
                return shape.CellExistsU[cellName, 0] != 0
                    ? shape.CellsU[cellName].ResultStr[""]
                    : null;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static void TrySetFormula(
            Visio.Shape shape, string cellName, string formula)
        {
            try
            {
                if (shape.CellExistsU[cellName, 0] != 0)
                    shape.CellsU[cellName].FormulaU = formula;
            }
            catch (COMException)
            {
                // The hidden layer is the primary suppression mechanism.
            }
        }

        private static void TryDelete(Visio.Shape shape)
        {
            try
            {
                shape?.Delete();
            }
            catch (COMException)
            {
                // The semantic stencil remains visible when a protected page
                // rejects cleanup, so the failed proxy cannot hide model data.
            }
        }
    }
}
