using alps.net.api.ALPS;
using alps.net.api.StandardPASS;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using static Microsoft.Office.Interop.Visio.VisRowTags;
using static Microsoft.Office.Interop.Visio.VisSectionIndices;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite
{
    public static class VisioHelper
    {
        public static void setVBAListenersRunning(Boolean newStatus)
        {
            Visio.IVDocument myActiveDocument = Globals.ThisAddIn.Application.ActiveDocument;

            if (myActiveDocument == null) return;

            SetBool(myActiveDocument.DocumentSheet, Constants.Properties.InteropWithVSTOShouldListenersRun, newStatus);
        }

        public enum VisioStencils
        {
            SID_STENCIL,
            SBD_STENCIL
        }

        /// <summary>
        /// Opens the latest SID-Stencil file from specified shape-folder
        /// </summary>
        /// <returns>The specified stencil file or null</returns>
        public static Visio.Document openStencil(VisioStencils stencil)
        {
            Visio.Documents visioDocs = Globals.ThisAddIn.Application.Documents;
            try
            {
                switch (stencil)
                {
                    case VisioStencils.SID_STENCIL:
                        Visio.Document sidShapes = visioDocs.OpenEx(ShapeFinder.getSIDName(),
                            (short)Visio.VisOpenSaveArgs.visOpenDocked);
                        return sidShapes;
                    case VisioStencils.SBD_STENCIL:
                        Visio.Document sbdShapes = visioDocs.OpenEx(ShapeFinder.getSBDName(),
                            (short)Visio.VisOpenSaveArgs.visOpenDocked);
                        return sbdShapes;
                }

            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                string msg = "Failed to load SID Shapes. Expecting file \"";
                switch (stencil)
                {
                    case VisioStencils.SID_STENCIL:
                        msg += ShapeFinder.getSIDName();
                        break;
                    case VisioStencils.SBD_STENCIL:
                        msg += ShapeFinder.getSBDName();
                        break;
                }
                msg += "\" to exist in \"my Shapes\" folder.\n";
                msg += "Error: " + e.Message;
                System.Windows.Forms.MessageBox.Show(msg);
            }
            return null;
        }

        public enum ShapeType
        {
            SBD, SID
        }

        public static Visio.Shape Place(string shapeType, Visio.Page page)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (string.IsNullOrWhiteSpace(shapeType)) throw new ArgumentException("A Visio master name is required.", nameof(shapeType));

            Visio.Document stencil = openStencil(GetStencil(shapeType));
            if (stencil == null) throw new InvalidOperationException("The required Visio stencil could not be opened.");

            Visio.Master sidMaster = stencil.Masters.get_ItemU(shapeType);

            Visio.Shape droppedShape = page.Drop(sidMaster, 0, 0);

            return droppedShape;
        }

        public static VisioStencils GetStencil(string shapeType)
        {
            if (SidMasterNames.Contains(shapeType)) return VisioStencils.SID_STENCIL;
            if (SbdMasterNames.Contains(shapeType)) return VisioStencils.SBD_STENCIL;

            throw new ArgumentException("Unknown ALPS Visio master: " + shapeType, nameof(shapeType));
        }

        private static readonly ISet<string> SidMasterNames = new HashSet<string>
        {
            Constants.SIDMasters.StandardActor,
            Constants.SIDMasters.InterfaceActor,
            Constants.SIDMasters.CommunicationRestriction,
            Constants.SIDMasters.StandardMessageConnector,
            Constants.SIDMasters.Message,
            Constants.SIDMasters.StandAloneMacro,
            Constants.SIDMasters.ActorExtension,
            Constants.SIDMasters.AbstractCommunicationChannel,
            Constants.SIDMasters.SystemInterfaceSubject,
            Constants.SIDMasters.SubjectGroup
        };

        private static readonly ISet<string> SbdMasterNames = new HashSet<string>
        {
            Constants.SBDMasters.DoState,
            Constants.SBDMasters.ReceiveState,
            Constants.SBDMasters.SendState,
            Constants.SBDMasters.GenericReturnToOriginReference,
            Constants.SBDMasters.StandardTransition,
            Constants.SBDMasters.ReceiveTransition,
            Constants.SBDMasters.SendingFailedTransition,
            Constants.SBDMasters.SendTransition,
            Constants.SBDMasters.FlowRestrictor,
            Constants.SBDMasters.TimeTransition,
            Constants.SBDMasters.UserCancelTransition
        };

        public static void SetHyperlink(Visio.Shape shape, string property, string value)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(property)) throw new ArgumentException("A hyperlink name is required.", nameof(property));
            EnsureSection(shape, visSectionHyperlink);
            if (shape.CellExistsU["Hyperlink." + property, 0] == 0)
                shape.AddNamedRow((short)visSectionHyperlink, property, (short)visTagDefault);

            shape.Hyperlinks.ItemU[property].SubAddress = value ?? string.Empty;
        }
        public static void SetProperty(Visio.Shape shape, string property, string value)
        {
            SetCell(shape, visSectionProp, property, CellFormulaMode.U, CellValueType.Literal, value);
        }
        public static void SetBool(Visio.Shape shape, string property, bool value)
        {
            SetCell(shape, visSectionProp, property, CellFormulaMode.U, CellValueType.Formula, value ? "TRUE" : "FALSE");
        }
        public static void SetSize(Visio.Shape shape, string cell, double value)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            shape.CellsU[cell].FormulaU = value.ToString(CultureInfo.InvariantCulture);
        }
        public static double GetSize(Visio.Shape shape, string cell)
        {
            return shape.CellsU[cell].Result[""];
        }

        public static void SetPropertyU(Visio.Shape shape, string property, object value)
        {
            SetCell(shape, visSectionProp, property, CellFormulaMode.U, CellValueType.Normal, value);
        }
        public static void SetPropertyULiteral(Visio.Shape shape, string property, object value)
        {
            SetCell(shape, visSectionProp, property, CellFormulaMode.U, CellValueType.Literal, value);
        }
        public static void SetPropertyFormulaU(Visio.Shape shape, string property, object value)
        {
            SetCell(shape, visSectionProp, property, CellFormulaMode.U, CellValueType.Formula, value);
        }
        public static void SetUser(Visio.Shape shape, string user, object value)
        {
            SetCell(shape, visSectionUser, user, CellFormulaMode.Normal, CellValueType.Literal, value);
        }
        public static void SetSizeMM(Visio.Shape shape, string cell, object value)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            string valueString = Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture) + " mm";
            shape.CellsU[cell].FormulaU = valueString;
        }
        private enum CellFormulaMode
        {
            Normal,
            U,
            Force,
            ForceU
        }
        private enum CellValueType
        {
            Literal,
            Formula,
            Size,
            Normal
        }
        private static void SetCell(Visio.Shape shape, Visio.VisSectionIndices? section, string rowName, CellFormulaMode formulaMode, CellValueType valueType, object value)
        {
            string sectionName = "";
            switch (section)
            {
                case visSectionProp: sectionName = "Prop."; break;
                case visSectionUser: sectionName = "User."; break;
            }

            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(rowName)) throw new ArgumentException("A cell name is required.", nameof(rowName));

            // Ensure row exists.
            if (section.HasValue) EnsureSection(shape, section.Value);
            if (shape.CellExistsU[sectionName + rowName, 0] == 0)
            {
                shape.AddNamedRow((short)section, rowName, (short)visTagDefault);
            }

            // Get the cell
            Visio.Cell cell = shape.CellsU[sectionName + rowName];

            // Convert value properly
            if (value is IFormattable formattable) value = formattable.ToString(null, CultureInfo.InvariantCulture);

            // Build the value string
            string valueString = "";
            switch (valueType)
            {
                case CellValueType.Formula: valueString = "=" + value; break;
                case CellValueType.Size: valueString = value + " mm"; break;
                case CellValueType.Literal: valueString = "\"" + EscapeFormulaLiteral(value) + "\""; break;
                case CellValueType.Normal: valueString = "" + value; break;
            }

            // Apply according to formula mode
            switch (formulaMode)
            {
                case CellFormulaMode.Normal: cell.Formula = valueString; break;
                case CellFormulaMode.U: cell.FormulaU = valueString; break;
                case CellFormulaMode.Force: cell.FormulaForce = valueString; break;
                case CellFormulaMode.ForceU: cell.FormulaForceU = valueString; break;
            }
        }

        private static string EscapeFormulaLiteral(object value)
        {
            return (value == null ? string.Empty : value.ToString()).Replace("\"", "\"\"");
        }

        private static void EnsureSection(Visio.Shape shape, Visio.VisSectionIndices section)
        {
            if (shape.SectionExists[(short)section, 0] == 0)
                shape.AddSection((short)section);
        }

        /// <summary>
        /// creates a new diagram page in visio
        /// and turns it into a sid page by setting all given parameters.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="nameU"></param>
        /// <param name="modelURI"></param>
        /// <param name="extends"></param>
        /// <param name="implements"></param>
        /// <param name="priority"></param>
        /// <returns>created visio page</returns>
        public static Visio.Page CreateSIDPage(string name, string nameU, string modelURI, string extends, string implements, string priority)
        {
            Visio.Application addin = Globals.ThisAddIn.Application;
            if (addin.Documents.Count < 1)
            {
                addin.Documents.Add("");
            }
            Visio.IVDocument document = Globals.ThisAddIn.Application.ActiveDocument;
            Visio.Page page = document.Pages.Add();

            string pageName = GetUniquePageName(document, name, "SID");
            string pageNameU = GetUniquePageName(document, nameU, pageName);
            page.Name = pageName;
            page.NameU = pageNameU;

            SetProperty(page.PageSheet, Constants.Properties.PageType, Constants.Properties.SIDPage);
            SetProperty(page.PageSheet, Constants.Properties.PageModelURI, modelURI);
            SetProperty(page.PageSheet, Constants.Properties.PageLayer, nameU);
            SetProperty(page.PageSheet, Constants.Properties.Transition.Extends, extends);
            SetProperty(page.PageSheet, Constants.Properties.Transition.Implements, implements);
            SetProperty(page.PageSheet, Constants.Properties.PriorityOrderNumber, priority);
            SetProperty(document.DocumentSheet, Constants.Properties.DocumentType, Constants.Properties.DocumentType);

            return page;
        }

        /// <summary>
        /// precondition: document and matching sid page already exist.
        /// </summary>
        /// <param name="sidPage"></param>
        /// <param name="name"></param>
        /// <param name="subjectShape">the subject the page belongs to</param>
        public static Visio.Page CreateSBDPage(Visio.Page sidPage, string name, string nameU, Visio.Shape subjectShape)
        {
            Debug.Print("creating new SBD page");
            if (sidPage == null) throw new ArgumentNullException(nameof(sidPage));
            if (subjectShape == null) throw new ArgumentNullException(nameof(subjectShape));

            Visio.IVDocument document = Globals.ThisAddIn.Application.ActiveDocument;
            Visio.Page page = document.Pages.Add();
            string pageName = GetUniquePageName(document, name, "SBD");
            string pageNameU = GetUniquePageName(document, nameU, pageName);
            page.Name = pageName;
            page.NameU = pageNameU;

            SetHyperlink(page.PageSheet, Constants.Properties.LinkedSIDPage, sidPage.NameU);
            SetProperty(page.PageSheet, Constants.Properties.PageLayer,
                sidPage.PageSheet.CellsU["Prop." + Constants.Properties.PageLayer].ResultStr[""]);
            SetProperty(page.PageSheet, Constants.Properties.PageType, Constants.Properties.SBDPage);
            SetProperty(page.PageSheet, Constants.Properties.SBDLinkedSubjectID, subjectShape.ID.ToString(CultureInfo.InvariantCulture));
            SetHyperlink(subjectShape, Constants.Properties.LinkedSBD, page.NameU);

            return page;
        }

        private static string GetUniquePageName(Visio.IVDocument document, string requestedName, string fallback)
        {
            string baseName = string.IsNullOrWhiteSpace(requestedName) ? fallback : requestedName.Trim();
            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Visio.Page existingPage in document.Pages)
            {
                usedNames.Add(existingPage.Name);
                usedNames.Add(existingPage.NameU);
            }

            if (!usedNames.Contains(baseName)) return baseName;

            for (int suffix = 2; ; suffix++)
            {
                string candidate = baseName + " (" + suffix.ToString(CultureInfo.InvariantCulture) + ")";
                if (!usedNames.Contains(candidate)) return candidate;
            }
        }

        public static List<ISimple2DVisualizationPoint> GetBounds(PASSProcessModelElement element)
        {
            if (OWLShapes.VisioLayout.TryGetGeneratedBounds(element, out List<ISimple2DVisualizationPoint> generatedBounds))
                return generatedBounds;

            List<ISimple2DVisualizationPoint> bounds = new List<ISimple2DVisualizationPoint>(
                element.getElementsWithUnspecifiedRelation().Values.OfType<ISimple2DVisualizationPoint>());

            if (bounds.Count < 2)
                throw new InvalidOperationException("Missing visualisation bounds for " + element.getModelComponentID() + ".");

            return bounds;
        }
    }
}
