using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using LayoutDirection = ALPS_Visio_AddIn_rewrite.VisioHelper.GraphLayoutDirection;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Rebinds connectors to deterministic shape sides after Visio moves nodes.
    /// </summary>
    internal static class VisioConnectorRebinder
    {
        public static void Rebind(Visio.IVPage page, LayoutDirection direction)
        {
            List<AutoArrangeConnector> connectors =
                new List<AutoArrangeConnector>();
            double pageCenterX =
                VisioShapeSheet.GetNumber(page.PageSheet, "PageWidth") / 2d;
            double pageCenterY =
                VisioShapeSheet.GetNumber(page.PageSheet, "PageHeight") / 2d;

            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.OneD == 0
                    || !TryGetConnectedShapes(page, shape,
                        out Visio.Shape source, out Visio.Shape target))
                {
                    continue;
                }

                AutoArrangeConnector connector =
                    new AutoArrangeConnector(shape, source, target);
                connectors.Add(connector);
            }

            int lowerSideCount = 0;
            int upperSideCount = 0;
            foreach (AutoArrangeConnector connector in connectors
                .OrderBy(candidate => candidate.Shape.ID))
            {
                AssignConnectionSides(connector, direction,
                    pageCenterX, pageCenterY,
                    ref lowerSideCount, ref upperSideCount);

                VisioRouting.TrySetCell(
                    connector.Shape, "ConFixedCode", 0, false);
                VisioRouting.TrySetCell(
                    connector.Shape, "ShapeRouteStyle", 1, false);
                VisioRouting.TrySetCell(
                    connector.Shape, "ConLineJumpCode", 0, false);
            }

            List<ConnectorEndpoint> endpoints = connectors
                .SelectMany(connector => new[]
                {
                    new ConnectorEndpoint(connector, true),
                    new ConnectorEndpoint(connector, false)
                })
                .ToList();

            foreach (IGrouping<string, ConnectorEndpoint> endpointGroup
                in endpoints.GroupBy(endpoint =>
                    endpoint.Shape.ID.ToString(CultureInfo.InvariantCulture)
                    + ":" + endpoint.Side))
            {
                List<ConnectorEndpoint> orderedEndpoints = endpointGroup
                    .OrderBy(GetOppositeShapePosition)
                    .ThenBy(endpoint => endpoint.Connector.Shape.ID)
                    .ToList();

                for (int index = 0; index < orderedEndpoints.Count; index++)
                {
                    GlueEndpointToSide(
                        orderedEndpoints[index],
                        GetPortPosition(index, orderedEndpoints.Count));
                }
            }
        }

        internal static bool TryGetConnectedShapes(Visio.IVPage page,
            Visio.Shape connector, out Visio.Shape source,
            out Visio.Shape target)
        {
            source = null;
            target = null;

            try
            {
                foreach (Visio.Connect connection in connector.Connects)
                {
                    int fromPart = connection.FromPart;
                    if (fromPart >= 7 && fromPart <= 9)
                        source = connection.ToSheet;
                    else if (fromPart >= 10 && fromPart <= 12)
                        target = connection.ToSheet;
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Recover missing endpoints from the semantic properties below.
            }

            if (TryGetStoredShape(page, connector,
                Constants.UserCells.AutoArrangeSourceShapeId,
                out Visio.Shape storedSource))
            {
                source = storedSource;
            }
            else if (TryGetRelatedShape(page, connector,
                Constants.Properties.MessageExchange.OriginSubject,
                out Visio.Shape semanticSource))
            {
                source = semanticSource;
            }
            else if (!IsNode(source))
            {
                source = null;
            }

            if (TryGetStoredShape(page, connector,
                Constants.UserCells.AutoArrangeTargetShapeId,
                out Visio.Shape storedTarget))
            {
                target = storedTarget;
            }
            else if (TryGetRelatedShape(page, connector,
                Constants.Properties.MessageExchange.TargetSubject,
                out Visio.Shape semanticTarget))
            {
                target = semanticTarget;
            }
            else if (!IsNode(target))
            {
                target = null;
            }

            return IsNode(source) && IsNode(target);
        }

        private static bool TryGetStoredShape(Visio.IVPage page,
            Visio.Shape connector, string userCell,
            out Visio.Shape storedShape)
        {
            storedShape = null;
            string shapeIdText = GetCellString(
                connector, "User." + userCell);
            if (!int.TryParse(shapeIdText, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int shapeId)
                || shapeId <= 0)
            {
                return false;
            }

            try
            {
                Visio.Shape candidate = page.Shapes.get_ItemFromID(shapeId);
                if (!IsNode(candidate)) return false;
                storedShape = candidate;
                return true;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private static bool TryGetRelatedShape(Visio.IVPage page,
            Visio.Shape connector, string relationProperty,
            out Visio.Shape relatedShape)
        {
            relatedShape = null;
            if (page == null || connector == null) return false;

            string relatedId = GetProperty(connector, relationProperty);
            if (string.IsNullOrWhiteSpace(relatedId) || relatedId == "0")
                return false;

            foreach (Visio.Shape candidate in page.Shapes)
            {
                if (!IsNode(candidate)) continue;

                string candidateId =
                    GetProperty(candidate, Constants.Properties.ID);
                if (IdentifiersMatch(relatedId, candidateId)
                    || IdentifiersMatch(
                        relatedId, GetShapeName(candidate, true))
                    || IdentifiersMatch(
                        relatedId, GetShapeName(candidate, false)))
                {
                    relatedShape = candidate;
                    return true;
                }
            }

            return false;
        }

        private static bool IsNode(Visio.Shape shape)
        {
            try
            {
                return shape != null && shape.OneD == 0;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private static string GetProperty(
            Visio.Shape shape, string propertyName)
        {
            return GetCellString(shape, "Prop." + propertyName);
        }

        private static string GetCellString(
            Visio.Shape shape, string cellName)
        {
            try
            {
                return shape.CellExistsU[cellName, 0] != 0
                    ? shape.CellsU[cellName].ResultStr[""]
                    : null;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        private static string GetShapeName(
            Visio.Shape shape, bool universal)
        {
            try
            {
                return universal ? shape.NameU : shape.Name;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        private static bool IdentifiersMatch(
            string expected, string candidate)
        {
            return !string.IsNullOrWhiteSpace(candidate)
                && string.Equals(expected.Trim(), candidate.Trim(),
                    StringComparison.OrdinalIgnoreCase);
        }

        private static void AssignConnectionSides(AutoArrangeConnector connector,
            LayoutDirection direction, double pageCenterX, double pageCenterY,
            ref int lowerSideCount, ref int upperSideCount)
        {
            double sourceX =
                VisioShapeSheet.GetNumber(connector.Source, "PinX");
            double sourceY =
                VisioShapeSheet.GetNumber(connector.Source, "PinY");
            double targetX =
                VisioShapeSheet.GetNumber(connector.Target, "PinX");
            double targetY =
                VisioShapeSheet.GetNumber(connector.Target, "PinY");
            const double sameRankTolerance = 0.05;

            if (connector.Source.ID == connector.Target.ID)
            {
                bool useLowerSide = lowerSideCount <= upperSideCount;
                ConnectionSide loopSide;
                if (direction == LayoutDirection.TopDown)
                    loopSide = useLowerSide
                        ? ConnectionSide.Left : ConnectionSide.Right;
                else
                    loopSide = useLowerSide
                        ? ConnectionSide.Bottom : ConnectionSide.Top;

                connector.SourceSide = loopSide;
                connector.TargetSide = loopSide;
                IncrementSideCount(
                    useLowerSide, ref lowerSideCount, ref upperSideCount);
                return;
            }

            if (direction == LayoutDirection.TopDown)
            {
                if (sourceY > targetY + sameRankTolerance)
                {
                    connector.SourceSide = ConnectionSide.Bottom;
                    connector.TargetSide = ConnectionSide.Top;
                }
                else if (sourceY < targetY - sameRankTolerance)
                {
                    bool useLeft = ChooseLowerSide(
                        (sourceX + targetX) / 2d, pageCenterX,
                        lowerSideCount, upperSideCount);
                    ConnectionSide feedbackSide = useLeft
                        ? ConnectionSide.Left : ConnectionSide.Right;
                    connector.SourceSide = feedbackSide;
                    connector.TargetSide = feedbackSide;
                    IncrementSideCount(
                        useLeft, ref lowerSideCount, ref upperSideCount);
                }
                else
                {
                    AssignHorizontalSides(connector, sourceX, targetX);
                }
            }
            else if (sourceX < targetX - sameRankTolerance)
            {
                connector.SourceSide = ConnectionSide.Right;
                connector.TargetSide = ConnectionSide.Left;
            }
            else if (sourceX > targetX + sameRankTolerance)
            {
                bool useBottom = ChooseLowerSide(
                    (sourceY + targetY) / 2d, pageCenterY,
                    lowerSideCount, upperSideCount);
                ConnectionSide feedbackSide = useBottom
                    ? ConnectionSide.Bottom : ConnectionSide.Top;
                connector.SourceSide = feedbackSide;
                connector.TargetSide = feedbackSide;
                IncrementSideCount(
                    useBottom, ref lowerSideCount, ref upperSideCount);
            }
            else
            {
                AssignVerticalSides(connector, sourceY, targetY);
            }
        }

        private static bool ChooseLowerSide(double midpoint,
            double pageCenter, int lowerSideCount, int upperSideCount)
        {
            const double centerTolerance = 0.25;
            if (midpoint < pageCenter - centerTolerance) return true;
            if (midpoint > pageCenter + centerTolerance) return false;
            return lowerSideCount <= upperSideCount;
        }

        private static void IncrementSideCount(bool useLowerSide,
            ref int lowerSideCount, ref int upperSideCount)
        {
            if (useLowerSide)
                lowerSideCount++;
            else
                upperSideCount++;
        }

        private static void AssignHorizontalSides(AutoArrangeConnector connector,
            double sourceX, double targetX)
        {
            bool targetIsRight = targetX >= sourceX;
            connector.SourceSide =
                targetIsRight ? ConnectionSide.Right : ConnectionSide.Left;
            connector.TargetSide =
                targetIsRight ? ConnectionSide.Left : ConnectionSide.Right;
        }

        private static void AssignVerticalSides(AutoArrangeConnector connector,
            double sourceY, double targetY)
        {
            bool targetIsAbove = targetY >= sourceY;
            connector.SourceSide =
                targetIsAbove ? ConnectionSide.Top : ConnectionSide.Bottom;
            connector.TargetSide =
                targetIsAbove ? ConnectionSide.Bottom : ConnectionSide.Top;
        }

        private static double GetOppositeShapePosition(ConnectorEndpoint endpoint)
        {
            Visio.Shape oppositeShape = endpoint.IsSource
                ? endpoint.Connector.Target
                : endpoint.Connector.Source;

            return endpoint.Side == ConnectionSide.Top
                || endpoint.Side == ConnectionSide.Bottom
                    ? VisioShapeSheet.GetNumber(oppositeShape, "PinX")
                    : VisioShapeSheet.GetNumber(oppositeShape, "PinY");
        }

        private static double GetPortPosition(int index, int count)
        {
            if (count <= 1)
                return 0.5;

            const double portMargin = 0.22;
            return portMargin
                + index * ((1d - 2d * portMargin) / (count - 1d));
        }

        private static void GlueEndpointToSide(ConnectorEndpoint endpoint,
            double portPosition)
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
            endpoint.Connector.Shape.CellsU[endpointCell]
                .GlueToPos(endpoint.Shape, x, y);
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
            public AutoArrangeConnector(Visio.Shape shape, Visio.Shape source,
                Visio.Shape target)
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
            public ConnectorEndpoint(
                AutoArrangeConnector connector, bool isSource)
            {
                Connector = connector;
                IsSource = isSource;
            }

            public AutoArrangeConnector Connector { get; private set; }
            public bool IsSource { get; private set; }
            public Visio.Shape Shape
            {
                get { return IsSource ? Connector.Source : Connector.Target; }
            }
            public ConnectionSide Side
            {
                get
                {
                    return IsSource
                        ? Connector.SourceSide
                        : Connector.TargetSide;
                }
            }
        }
    }
}
