using System;
using System.Globalization;
using Visio = Microsoft.Office.Interop.Visio;
using LayoutDirection = ALPS_Visio_AddIn_rewrite.VisioHelper.GraphLayoutDirection;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Applies Visio's connected-graph layout and deterministically rebinds
    /// connector endpoints after node placement.
    /// </summary>
    internal static class VisioAutoArrange
    {
        public static void Arrange(Visio.IVPage page, LayoutDirection direction)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (!Enum.IsDefined(typeof(LayoutDirection), direction))
                throw new ArgumentOutOfRangeException(nameof(direction));

            page.PageSheet.CellsU["PlaceStyle"].FormulaU =
                ((int)direction).ToString(CultureInfo.InvariantCulture);
            ConfigureSpacing(page.PageSheet, direction);
            VisioRouting.TrySetCell(page.PageSheet, "RouteStyle", 1, false);

            AlpsDiagramKind diagramKind = GetDiagramKind(page);
            if (diagramKind == AlpsDiagramKind.Sid)
                VisioSidMessageConnectorRenderer.Ensure(page);

            foreach (Visio.Shape shape in page.Shapes)
            {
                if (VisioSidMessageConnectorRenderer.IsSemanticShadow(shape)
                    || VisioSidMessageConnectorRenderer.IsVisualConnector(shape)
                    || shape.OneD == 0
                    || shape.CellExistsU["ConFixedCode", 0] == 0)
                {
                    continue;
                }
                shape.CellsU["ConFixedCode"].FormulaU = "0";
            }

            bool arranged = VisioGraphAutoArranger.TryArrange(
                page, direction, diagramKind);
            if (!arranged)
                page.Layout();

            ConfigurePrintLayout(page.PageSheet, direction);
            if (diagramKind == AlpsDiagramKind.Sid)
            {
                // The SID master couples a Message Box to connector control
                // cells. Moving the box can change the connector geometry and
                // invalidate Glue entries, so no box may move after the final
                // endpoint binding.
                VisioMessageContainerPositioner.Reposition(page, direction);
            }
            VisioConnectorRebinder.Rebind(page, direction);
            if (diagramKind == AlpsDiagramKind.Sid)
            {
                // Visio collapses two same-side endpoints onto the subject
                // border when their coordinates align. Rebuild SID channels
                // last so their explicit message corridors cannot be rerouted
                // by a later layout or endpoint-binding operation.
                VisioSidMessageConnectorRenderer
                    .RouteAlongMessageCorridors(page, direction);
            }
        }

        private static void ConfigureSpacing(Visio.Shape pageSheet,
            LayoutDirection direction)
        {
            VisioRouting.TrySetCell(pageSheet, "EnableGrid", 1, false);
            VisioRouting.TrySetCell(pageSheet, "ResizePage", 1, false);
            VisioRouting.TrySetCell(pageSheet, "PlaceDepth", 2, false);
            VisioRouting.TrySetCell(pageSheet, "BlockSizeX", 90, true);
            VisioRouting.TrySetCell(pageSheet, "BlockSizeY", 30, true);

            if (direction == LayoutDirection.TopDown)
            {
                VisioRouting.TrySetCell(pageSheet, "AvenueSizeX", 45, true);
                VisioRouting.TrySetCell(pageSheet, "AvenueSizeY", 35, true);
            }
            else
            {
                VisioRouting.TrySetCell(pageSheet, "AvenueSizeX", 45, true);
                VisioRouting.TrySetCell(pageSheet, "AvenueSizeY", 30, true);
            }

            VisioRouting.TrySetCell(pageSheet, "LineToNodeX", 12, true);
            VisioRouting.TrySetCell(pageSheet, "LineToNodeY", 12, true);
            VisioRouting.TrySetCell(pageSheet, "LineToLineX", 8, true);
            VisioRouting.TrySetCell(pageSheet, "LineToLineY", 8, true);
        }

        private static void ConfigurePrintLayout(
            Visio.Shape pageSheet, LayoutDirection direction)
        {
            VisioShapeSheet.SetNumber(
                pageSheet, "PrintPageOrientation",
                direction == LayoutDirection.LeftRight ? 2d : 1d);
            VisioShapeSheet.SetNumber(pageSheet, "OnPage", 1d);
            VisioShapeSheet.SetNumber(pageSheet, "PagesX", 1d);
            VisioShapeSheet.SetNumber(pageSheet, "PagesY", 1d);
        }

        private static AlpsDiagramKind GetDiagramKind(Visio.IVPage page)
        {
            string pageTypeCell = "Prop." + Constants.Properties.PageType;
            try
            {
                if (page.PageSheet.CellExistsU[pageTypeCell, 0] == 0)
                    return InferDiagramKind(page);

                string pageType =
                    page.PageSheet.CellsU[pageTypeCell].ResultStr[""];
                if (string.Equals(pageType, Constants.Properties.SIDPage,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return AlpsDiagramKind.Sid;
                }

                if (string.Equals(pageType, Constants.Properties.SBDPage,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return AlpsDiagramKind.Sbd;
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return InferDiagramKind(page);
            }

            return InferDiagramKind(page);
        }

        private static AlpsDiagramKind InferDiagramKind(Visio.IVPage page)
        {
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.OneD != 0) continue;
                try
                {
                    if (shape.HasCategory(
                        Constants.ShapeCategories.SBDState))
                    {
                        return AlpsDiagramKind.Sbd;
                    }

                    if (shape.HasCategory(
                            Constants.ShapeCategories.SIDSubject)
                        || shape.HasCategory(
                            Constants.ShapeCategories.SIDSubjectWithSBD))
                    {
                        return AlpsDiagramKind.Sid;
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Try the remaining shapes before using Visio's fallback.
                }
            }

            return AlpsDiagramKind.Unknown;
        }
    }
}
