using System;
using System.Globalization;
using static Microsoft.Office.Interop.Visio.VisRowTags;
using static Microsoft.Office.Interop.Visio.VisSectionIndices;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Provides a single boundary for reading and writing Visio ShapeSheet cells.
    /// </summary>
    internal static class VisioShapeSheet
    {
        public static void SetHyperlink(Visio.Shape shape, string rowName, string subAddress)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(rowName))
                throw new ArgumentException("A hyperlink name is required.", nameof(rowName));

            EnsureSection(shape, visSectionHyperlink);
            if (shape.CellExistsU["Hyperlink." + rowName, 0] == 0)
                shape.AddNamedRow((short)visSectionHyperlink, rowName, (short)visTagDefault);

            shape.Hyperlinks.ItemU[rowName].SubAddress = subAddress ?? string.Empty;
        }

        public static void SetProperty(Visio.Shape shape, string rowName, string value)
        {
            SetCell(shape, visSectionProp, rowName, FormulaMode.Universal, ValueType.Literal, value);
        }

        public static void SetBooleanProperty(Visio.Shape shape, string rowName, bool value)
        {
            SetCell(shape, visSectionProp, rowName, FormulaMode.Universal, ValueType.Formula,
                value ? "TRUE" : "FALSE");
        }

        public static void SetUniversalProperty(Visio.Shape shape, string rowName, object value)
        {
            SetCell(shape, visSectionProp, rowName, FormulaMode.Universal, ValueType.Raw, value);
        }

        public static void SetUniversalLiteralProperty(Visio.Shape shape, string rowName, object value)
        {
            SetCell(shape, visSectionProp, rowName, FormulaMode.Universal, ValueType.Literal, value);
        }

        public static void SetUniversalFormulaProperty(Visio.Shape shape, string rowName, object value)
        {
            SetCell(shape, visSectionProp, rowName, FormulaMode.Universal, ValueType.Formula, value);
        }

        public static void SetUserCell(Visio.Shape shape, string rowName, object value)
        {
            SetCell(shape, visSectionUser, rowName, FormulaMode.Local, ValueType.Literal, value);
        }

        public static void SetNumber(Visio.Shape shape, string cellName, double value)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(cellName))
                throw new ArgumentException("A cell name is required.", nameof(cellName));

            shape.CellsU[cellName].FormulaU = value.ToString(CultureInfo.InvariantCulture);
        }

        public static void SetMillimeters(Visio.Shape shape, string cellName, object value)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(cellName))
                throw new ArgumentException("A cell name is required.", nameof(cellName));

            double numericValue = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            shape.CellsU[cellName].FormulaU =
                numericValue.ToString(CultureInfo.InvariantCulture) + " mm";
        }

        public static double GetNumber(Visio.Shape shape, string cellName)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(cellName))
                throw new ArgumentException("A cell name is required.", nameof(cellName));

            return shape.CellsU[cellName].Result[""];
        }

        private static void SetCell(Visio.Shape shape, Visio.VisSectionIndices section,
            string rowName, FormulaMode formulaMode, ValueType valueType, object value)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(rowName))
                throw new ArgumentException("A cell name is required.", nameof(rowName));

            EnsureSection(shape, section);
            string cellName = GetSectionPrefix(section) + rowName;
            if (shape.CellExistsU[cellName, 0] == 0)
                shape.AddNamedRow((short)section, rowName, (short)visTagDefault);

            string formula = FormatValue(value, valueType);
            Visio.Cell cell = shape.CellsU[cellName];
            if (formulaMode == FormulaMode.Universal)
                cell.FormulaU = formula;
            else
                cell.Formula = formula;
        }

        private static string GetSectionPrefix(Visio.VisSectionIndices section)
        {
            switch (section)
            {
                case visSectionProp:
                    return "Prop.";
                case visSectionUser:
                    return "User.";
                default:
                    throw new ArgumentOutOfRangeException(nameof(section));
            }
        }

        private static string FormatValue(object value, ValueType valueType)
        {
            string invariantValue = value is IFormattable formattable
                ? formattable.ToString(null, CultureInfo.InvariantCulture)
                : value?.ToString();
            invariantValue = invariantValue ?? string.Empty;

            switch (valueType)
            {
                case ValueType.Literal:
                    return "\"" + invariantValue.Replace("\"", "\"\"") + "\"";
                case ValueType.Formula:
                    return "=" + invariantValue;
                case ValueType.Raw:
                    return invariantValue;
                default:
                    throw new ArgumentOutOfRangeException(nameof(valueType));
            }
        }

        private static void EnsureSection(Visio.Shape shape, Visio.VisSectionIndices section)
        {
            if (shape.SectionExists[(short)section, 0] == 0)
                shape.AddSection((short)section);
        }

        private enum FormulaMode
        {
            Local,
            Universal
        }

        private enum ValueType
        {
            Literal,
            Formula,
            Raw
        }
    }
}
