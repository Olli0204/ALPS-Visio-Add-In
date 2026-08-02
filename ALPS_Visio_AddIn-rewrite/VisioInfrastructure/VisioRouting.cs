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
        private const string SolidLinePatternFormula = "1";
        private const string StandardEndArrowFormula = "4";
        private const string MediumArrowSizeFormula = "2";

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

        /// <summary>
        /// Keeps imported directional connectors visible even when a stencil
        /// references a custom line pattern or line end that is absent from the
        /// destination document stencil.
        /// </summary>
        public static void EnsureImportedDirectionalLine(
            Visio.Shape connector)
        {
            if (connector == null)
                throw new ArgumentNullException(nameof(connector));

            bool arrowFallbackApplied = EnsureBuiltInCell(
                connector, "EndArrow", StandardEndArrowFormula,
                RequiresArrowFallback);
            EnsureBuiltInCell(
                connector, "LinePattern", SolidLinePatternFormula,
                RequiresPatternFallback);

            if (arrowFallbackApplied)
                TrySetFormulaForceU(connector, "EndArrowSize",
                    MediumArrowSizeFormula);
        }

        internal static bool RequiresPatternFallback(
            double value, string formula)
        {
            return value < 1d || value > 23d || UsesCustomResource(formula);
        }

        internal static bool RequiresArrowFallback(
            double value, string formula)
        {
            return value < 1d || value > 45d || UsesCustomResource(formula);
        }

        private static bool UsesCustomResource(string formula)
        {
            return !string.IsNullOrWhiteSpace(formula)
                && formula.IndexOf(
                    "USE(", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool EnsureBuiltInCell(Visio.Shape shape,
            string cellName, string fallbackFormula,
            Func<double, string, bool> requiresFallback)
        {
            try
            {
                if (shape.CellExistsU[cellName, 0] == 0)
                    return false;

                Visio.Cell cell = shape.CellsU[cellName];
                string formula;
                try
                {
                    formula = cell.FormulaU;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    formula = null;
                }

                if (UsesCustomResource(formula))
                {
                    cell.FormulaForceU = fallbackFormula;
                    return true;
                }

                double value;
                try
                {
                    value = cell.Result[""];
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    cell.FormulaForceU = fallbackFormula;
                    return true;
                }

                if (!requiresFallback(value, formula))
                    return false;

                cell.FormulaForceU = fallbackFormula;
                return true;
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                Debug.Print("Could not normalize Visio line cell "
                    + cellName + ": " + exception.Message);
                return false;
            }
        }

        private static bool TrySetFormulaForceU(Visio.Shape shape,
            string cellName, string formula)
        {
            try
            {
                if (shape.CellExistsU[cellName, 0] == 0)
                    return false;

                shape.CellsU[cellName].FormulaForceU = formula;
                return true;
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                Debug.Print("Could not set Visio line cell " + cellName
                    + ": " + exception.Message);
                return false;
            }
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
