using alps.net.api.ALPS;
using alps.net.api.StandardPASS;
using System;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    /// <summary>
    /// Supplies deterministic fallback geometry for model elements that do not
    /// contain ALPS 2D visualisation data.
    /// </summary>
    internal static class VisioLayout
    {
        private const int SidColumns = 3;
        private const int SbdColumns = 4;

        public static void PrepareOrArrange(IVisioExportableWithShape exportable, int index, bool subjectDiagram)
        {
            if (exportable == null || exportable.PrepareDimensions()) return;

            IPASSProcessModelElement element = exportable as IPASSProcessModelElement;
            if (element == null) return;

            int columns = subjectDiagram ? SbdColumns : SidColumns;
            int column = index % columns;
            int row = index / columns;
            double width = subjectDiagram ? 0.18 : 0.22;
            double height = subjectDiagram ? 0.12 : 0.16;
            double x = 0.12 + column * ((1.0 - 0.24) / Math.Max(1, columns - 1));
            double y = Math.Max(0.12, 0.82 - row * 0.20);

            AddPoint(element, x, y);
            AddPoint(element, width, height);
        }

        private static void AddPoint(IPASSProcessModelElement element, double x, double y)
        {
            Simple2DVisualizationPoint point = new Simple2DVisualizationPoint();
            point.setRelative2DPosX(x);
            point.setRelative2DPosY(y);
            element.addElementWithUnspecifiedRelation(point);
        }
    }
}
