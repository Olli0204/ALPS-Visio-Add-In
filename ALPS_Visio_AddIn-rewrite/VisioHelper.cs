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

        private static readonly IDictionary<VisioStencils, Visio.Document> OpenStencils =
            new Dictionary<VisioStencils, Visio.Document>();

        /// <summary>
        /// Opens the latest SID-Stencil file from specified shape-folder
        /// </summary>
        /// <returns>The specified stencil file or null</returns>
        public static Visio.Document openStencil(VisioStencils stencil)
        {
            return OpenStencil(stencil, false);
        }

        /// <summary>
        /// Opens both master sources once for an OWL import. Import-specific drop
        /// behavior is implemented by the C# exporter, so neither stencil needs
        /// to execute VBA or display a macro security prompt.
        /// </summary>
        public static void OpenImportStencils()
        {
            OpenStencil(VisioStencils.SID_STENCIL, true);
            OpenStencil(VisioStencils.SBD_STENCIL, true);
        }

        private static Visio.Document OpenStencil(VisioStencils stencil, bool disableMacros)
        {
            Visio.Documents visioDocs = Globals.ThisAddIn.Application.Documents;
            string stencilName = GetStencilName(stencil);
            Visio.Document openDocument = FindOpenStencil(visioDocs, stencil, stencilName);
            if (openDocument != null) return openDocument;

            try
            {
                int flags = (int)Visio.VisOpenSaveArgs.visOpenDocked;
                if (disableMacros)
                    flags |= (int)Visio.VisOpenSaveArgs.visOpenMacrosDisabled;

                Visio.Document document = visioDocs.OpenEx(stencilName, (short)flags);
                OpenStencils[stencil] = document;
                return document;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                string msg = "Failed to load ALPS/PASS shapes. Expecting file \""
                    + stencilName + "\" to exist in \"My Shapes\".\n";
                msg += "Error: " + e.Message;
                System.Windows.Forms.MessageBox.Show(msg);
            }
            return null;
        }

        private static string GetStencilName(VisioStencils stencil)
        {
            switch (stencil)
            {
                case VisioStencils.SID_STENCIL:
                    return ShapeFinder.getSIDName();
                case VisioStencils.SBD_STENCIL:
                    return ShapeFinder.getSBDName();
                default:
                    throw new ArgumentOutOfRangeException(nameof(stencil));
            }
        }

        private static Visio.Document FindOpenStencil(Visio.Documents documents,
            VisioStencils stencil, string stencilName)
        {
            if (OpenStencils.TryGetValue(stencil, out Visio.Document cachedDocument))
            {
                try
                {
                    if (cachedDocument != null
                        && string.Equals(cachedDocument.Name, stencilName, StringComparison.OrdinalIgnoreCase))
                        return cachedDocument;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // The user closed the cached stencil. Search the live
                    // Documents collection before opening it again.
                }

                OpenStencils.Remove(stencil);
            }

            foreach (Visio.Document document in documents)
            {
                try
                {
                    if (!string.Equals(document.Name, stencilName, StringComparison.OrdinalIgnoreCase)) continue;
                    OpenStencils[stencil] = document;
                    return document;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Ignore a document that is currently being closed.
                }
            }

            return null;
        }

        public enum ShapeType
        {
            SBD, SID
        }

        public enum GraphLayoutDirection
        {
            TopDown = 1,
            LeftRight = 2
        }

        /// <summary>
        /// Applies Visio's connected-graph layout to a drawing page. Fixed
        /// imported connectors are released first so they can be rerouted around
        /// the newly positioned nodes.
        /// </summary>
        public static void AutoArrangePage(Visio.IVPage page, GraphLayoutDirection direction)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (!Enum.IsDefined(typeof(GraphLayoutDirection), direction))
                throw new ArgumentOutOfRangeException(nameof(direction));

            page.PageSheet.CellsU["PlaceStyle"].FormulaU =
                ((int)direction).ToString(CultureInfo.InvariantCulture);
            ConfigureAutoArrangeSpacing(page.PageSheet, direction);
            TrySetRoutingCell(page.PageSheet, "RouteStyle", 1, false);

            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.OneD == 0 || shape.CellExistsU["ConFixedCode", 0] == 0) continue;
                shape.CellsU["ConFixedCode"].FormulaU = "0";
            }

            page.Layout();
            RebindAutoArrangeConnectors(page, direction);
        }

        private static void ConfigureAutoArrangeSpacing(Visio.Shape pageSheet,
            GraphLayoutDirection direction)
        {
            // The ALPS state masters are considerably larger than Visio's
            // default flowchart shapes. A larger internal grid prevents state
            // labels, transition labels and parallel branches from competing
            // for the same narrow routing lanes.
            TrySetRoutingCell(pageSheet, "EnableGrid", 1, false);
            TrySetRoutingCell(pageSheet, "ResizePage", 1, false);
            TrySetRoutingCell(pageSheet, "PlaceDepth", 2, false);
            TrySetRoutingCell(pageSheet, "BlockSizeX", 90, true);
            TrySetRoutingCell(pageSheet, "BlockSizeY", 30, true);

            if (direction == GraphLayoutDirection.TopDown)
            {
                TrySetRoutingCell(pageSheet, "AvenueSizeX", 45, true);
                TrySetRoutingCell(pageSheet, "AvenueSizeY", 35, true);
            }
            else
            {
                TrySetRoutingCell(pageSheet, "AvenueSizeX", 45, true);
                TrySetRoutingCell(pageSheet, "AvenueSizeY", 30, true);
            }

            TrySetRoutingCell(pageSheet, "LineToNodeX", 12, true);
            TrySetRoutingCell(pageSheet, "LineToNodeY", 12, true);
            TrySetRoutingCell(pageSheet, "LineToLineX", 8, true);
            TrySetRoutingCell(pageSheet, "LineToLineY", 8, true);
        }

        private static void RebindAutoArrangeConnectors(Visio.IVPage page,
            GraphLayoutDirection direction)
        {
            List<AutoArrangeConnector> connectors = new List<AutoArrangeConnector>();
            double pageCenterX = GetSize(page.PageSheet, "PageWidth") / 2d;
            double pageCenterY = GetSize(page.PageSheet, "PageHeight") / 2d;

            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.OneD == 0
                    || !TryGetConnectedShapes(shape, out Visio.Shape source, out Visio.Shape target))
                    continue;

                AutoArrangeConnector connector = new AutoArrangeConnector(shape, source, target);
                AssignConnectionSides(connector, direction, pageCenterX, pageCenterY);
                connectors.Add(connector);

                TrySetRoutingCell(shape, "ConFixedCode", 0, false);
                TrySetRoutingCell(shape, "ShapeRouteStyle", 1, false);
            }

            List<ConnectorEndpoint> endpoints = connectors
                .SelectMany(connector => new[]
                {
                    new ConnectorEndpoint(connector, true),
                    new ConnectorEndpoint(connector, false)
                })
                .ToList();

            foreach (IGrouping<string, ConnectorEndpoint> endpointGroup in endpoints.GroupBy(endpoint =>
                endpoint.Shape.ID.ToString(CultureInfo.InvariantCulture) + ":" + endpoint.Side))
            {
                List<ConnectorEndpoint> orderedEndpoints = endpointGroup
                    .OrderBy(GetOppositeShapePosition)
                    .ThenBy(endpoint => endpoint.Connector.Shape.ID)
                    .ToList();

                for (int index = 0; index < orderedEndpoints.Count; index++)
                {
                    double portPosition = GetAutoArrangePortPosition(index, orderedEndpoints.Count);
                    GlueEndpointToSide(orderedEndpoints[index], portPosition);
                }
            }
        }

        private static bool TryGetConnectedShapes(Visio.Shape connector,
            out Visio.Shape source, out Visio.Shape target)
        {
            source = null;
            target = null;

            foreach (Visio.Connect connection in connector.Connects)
            {
                int fromPart = connection.FromPart;
                if (fromPart >= 7 && fromPart <= 9)
                    source = connection.ToSheet;
                else if (fromPart >= 10 && fromPart <= 12)
                    target = connection.ToSheet;
            }

            return source != null && target != null && source.OneD == 0 && target.OneD == 0;
        }

        private static void AssignConnectionSides(AutoArrangeConnector connector,
            GraphLayoutDirection direction, double pageCenterX, double pageCenterY)
        {
            double sourceX = GetSize(connector.Source, "PinX");
            double sourceY = GetSize(connector.Source, "PinY");
            double targetX = GetSize(connector.Target, "PinX");
            double targetY = GetSize(connector.Target, "PinY");
            const double sameRankTolerance = 0.05;

            if (connector.Source.ID == connector.Target.ID)
            {
                ConnectionSide loopSide = direction == GraphLayoutDirection.TopDown
                    ? ConnectionSide.Right
                    : ConnectionSide.Top;
                connector.SourceSide = loopSide;
                connector.TargetSide = loopSide;
                return;
            }

            if (direction == GraphLayoutDirection.TopDown)
            {
                if (sourceY > targetY + sameRankTolerance)
                {
                    connector.SourceSide = ConnectionSide.Bottom;
                    connector.TargetSide = ConnectionSide.Top;
                }
                else if (sourceY < targetY - sameRankTolerance)
                {
                    ConnectionSide feedbackSide = (sourceX + targetX) / 2d <= pageCenterX
                        ? ConnectionSide.Left
                        : ConnectionSide.Right;
                    connector.SourceSide = feedbackSide;
                    connector.TargetSide = feedbackSide;
                }
                else
                {
                    AssignHorizontalSides(connector, sourceX, targetX);
                }
            }
            else
            {
                if (sourceX < targetX - sameRankTolerance)
                {
                    connector.SourceSide = ConnectionSide.Right;
                    connector.TargetSide = ConnectionSide.Left;
                }
                else if (sourceX > targetX + sameRankTolerance)
                {
                    ConnectionSide feedbackSide = (sourceY + targetY) / 2d <= pageCenterY
                        ? ConnectionSide.Bottom
                        : ConnectionSide.Top;
                    connector.SourceSide = feedbackSide;
                    connector.TargetSide = feedbackSide;
                }
                else
                {
                    AssignVerticalSides(connector, sourceY, targetY);
                }
            }
        }

        private static void AssignHorizontalSides(AutoArrangeConnector connector,
            double sourceX, double targetX)
        {
            bool targetIsRight = targetX >= sourceX;
            connector.SourceSide = targetIsRight ? ConnectionSide.Right : ConnectionSide.Left;
            connector.TargetSide = targetIsRight ? ConnectionSide.Left : ConnectionSide.Right;
        }

        private static void AssignVerticalSides(AutoArrangeConnector connector,
            double sourceY, double targetY)
        {
            bool targetIsAbove = targetY >= sourceY;
            connector.SourceSide = targetIsAbove ? ConnectionSide.Top : ConnectionSide.Bottom;
            connector.TargetSide = targetIsAbove ? ConnectionSide.Bottom : ConnectionSide.Top;
        }

        private static double GetOppositeShapePosition(ConnectorEndpoint endpoint)
        {
            Visio.Shape oppositeShape = endpoint.IsSource
                ? endpoint.Connector.Target
                : endpoint.Connector.Source;

            return endpoint.Side == ConnectionSide.Top || endpoint.Side == ConnectionSide.Bottom
                ? GetSize(oppositeShape, "PinX")
                : GetSize(oppositeShape, "PinY");
        }

        private static double GetAutoArrangePortPosition(int index, int count)
        {
            if (count <= 1) return 0.5;
            const double portMargin = 0.22;
            return portMargin + index * ((1d - 2d * portMargin) / (count - 1d));
        }

        private static void GlueEndpointToSide(ConnectorEndpoint endpoint, double portPosition)
        {
            double x;
            double y;
            switch (endpoint.Side)
            {
                case ConnectionSide.Left:
                    x = 0;
                    y = portPosition;
                    break;
                case ConnectionSide.Right:
                    x = 1;
                    y = portPosition;
                    break;
                case ConnectionSide.Bottom:
                    x = portPosition;
                    y = 0;
                    break;
                case ConnectionSide.Top:
                    x = portPosition;
                    y = 1;
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            string endpointCell = endpoint.IsSource ? "BeginX" : "EndX";
            endpoint.Connector.Shape.CellsU[endpointCell].GlueToPos(endpoint.Shape, x, y);
        }

        private enum ConnectionSide
        {
            Left,
            Right,
            Bottom,
            Top
        }

        private sealed class AutoArrangeConnector
        {
            public AutoArrangeConnector(Visio.Shape shape, Visio.Shape source, Visio.Shape target)
            {
                Shape = shape;
                Source = source;
                Target = target;
            }

            public Visio.Shape Shape { get; private set; }
            public Visio.Shape Source { get; private set; }
            public Visio.Shape Target { get; private set; }
            public ConnectionSide SourceSide { get; set; }
            public ConnectionSide TargetSide { get; set; }
        }

        private sealed class ConnectorEndpoint
        {
            public ConnectorEndpoint(AutoArrangeConnector connector, bool isSource)
            {
                Connector = connector;
                IsSource = isSource;
            }

            public AutoArrangeConnector Connector { get; private set; }
            public bool IsSource { get; private set; }
            public Visio.Shape Shape { get { return IsSource ? Connector.Source : Connector.Target; } }
            public ConnectionSide Side { get { return IsSource ? Connector.SourceSide : Connector.TargetSide; } }
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
            Constants.SIDMasters.MessageBox,
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

        /// <summary>
        /// Configures routing only. Node placement remains controlled by
        /// <see cref="OWLShapes.VisioLayout"/> and is never handed to Visio's
        /// automatic page layout.
        /// </summary>
        public static void ConfigureFallbackSbdRouting(Visio.Page page)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));

            TrySetRoutingCell(page.PageSheet, "RouteStyle", 1, false); // visLORouteRightAngle
            TrySetRoutingCell(page.PageSheet, "LineToLineX", 10, true);
            TrySetRoutingCell(page.PageSheet, "LineToLineY", 10, true);
        }

        /// <summary>
        /// Routes regular transitions with one deterministic horizontal/vertical
        /// bend between their ordered ports. Feedback transitions use the
        /// obstacle-avoiding right-angle router outside the state graph.
        /// </summary>
        public static void ConfigureFallbackTransitionRouting(Visio.Shape connector, bool isFeedback)
        {
            if (connector == null) throw new ArgumentNullException(nameof(connector));

            TrySetRoutingCell(connector, "ShapeRouteStyle", isFeedback ? 1 : 21, false);
            TrySetRoutingCell(connector, "ConFixedCode", 0, false);
            TrySetRoutingCell(connector, "ConLineJumpCode", 0, false);
        }

        public static void FinalizeFallbackTransitionRouting(Visio.Shape connector, bool isFeedback)
        {
            if (connector == null) throw new ArgumentNullException(nameof(connector));

            // Freeze the calculated single-bend route. Feedback connectors
            // remain reroutable so Visio can keep them outside moved states.
            TrySetRoutingCell(connector, "ConFixedCode", isFeedback ? 1 : 2, false);
        }

        private static bool TrySetRoutingCell(Visio.Shape shape, string cell, double value, bool millimeters)
        {
            try
            {
                if (shape.CellExistsU[cell, 0] == 0)
                {
                    Debug.Print("Visio routing cell is unavailable: " + cell);
                    return false;
                }

                if (millimeters) SetSizeMM(shape, cell, value);
                else SetSize(shape, cell, value);
                return true;
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                // Routing hints are optional. Older Visio versions and some
                // custom masters do not expose every ShapeSheet routing cell.
                Debug.Print("Could not set Visio routing cell " + cell + ": " + exception.Message);
                return false;
            }
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
