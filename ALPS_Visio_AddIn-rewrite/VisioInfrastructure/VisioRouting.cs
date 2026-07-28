using System;
using System.Diagnostics;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Applies optional routing hints to Visio pages and connectors.
    /// </summary>
    internal static class VisioRouting
    {
        public static void ConfigureFallbackSbd(Visio.Page page)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));

            TrySetCell(page.PageSheet, "RouteStyle", 1, false);
            TrySetCell(page.PageSheet, "LineToLineX", 10, true);
            TrySetCell(page.PageSheet, "LineToLineY", 10, true);
        }

        public static void ConfigureFallbackTransition(Visio.Shape connector,
            bool isFeedback)
        {
            if (connector == null) throw new ArgumentNullException(nameof(connector));

            TrySetCell(connector, "ShapeRouteStyle", isFeedback ? 1 : 21, false);
            TrySetCell(connector, "ConFixedCode", 0, false);
            TrySetCell(connector, "ConLineJumpCode", 0, false);
        }

        public static void FinalizeFallbackTransition(Visio.Shape connector,
            bool isFeedback)
        {
            if (connector == null) throw new ArgumentNullException(nameof(connector));

            TrySetCell(connector, "ConFixedCode", isFeedback ? 1 : 2, false);
        }

        public static bool TrySetCell(Visio.Shape shape, string cellName, double value,
            bool millimeters)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(cellName))
                throw new ArgumentException("A routing cell name is required.", nameof(cellName));

            try
            {
                if (shape.CellExistsU[cellName, 0] == 0)
                {
                    Debug.Print("Visio routing cell is unavailable: " + cellName);
                    return false;
                }

                if (millimeters)
                    VisioShapeSheet.SetMillimeters(shape, cellName, value);
                else
                    VisioShapeSheet.SetNumber(shape, cellName, value);

                return true;
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                // Routing hints are optional. Older Visio versions and custom
                // masters do not expose every ShapeSheet routing cell.
                Debug.Print("Could not set Visio routing cell " + cellName + ": "
                    + exception.Message);
                return false;
            }
        }
    }
}
