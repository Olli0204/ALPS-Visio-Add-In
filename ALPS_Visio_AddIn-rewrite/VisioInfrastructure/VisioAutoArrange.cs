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

            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.OneD == 0 || shape.CellExistsU["ConFixedCode", 0] == 0)
                    continue;
                shape.CellsU["ConFixedCode"].FormulaU = "0";
            }

            AlpsDiagramKind diagramKind = GetDiagramKind(page);
            bool arranged = VisioGraphAutoArranger.TryArrange(
                page, direction, diagramKind);
            if (!arranged)
                page.Layout();

            VisioConnectorRebinder.Rebind(page, direction);
            if (diagramKind == AlpsDiagramKind.Sid)
                VisioMessageContainerPositioner.Reposition(page, direction);
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
