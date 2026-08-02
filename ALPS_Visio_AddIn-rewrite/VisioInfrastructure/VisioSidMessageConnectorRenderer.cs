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
        private const double CorridorClearance = 0.55;
        private const double CoordinateTolerance = 0.05;

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

        public static void Ensure(Visio.IVPage page)
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
                if (visualConnector == null)
                {
                    visualConnector = CreateVisualConnector(
                        page, semanticConnector);
                }

                if (visualConnector == null) continue;

                MarkVisualConnector(
                    visualConnector, semanticConnector, source, target);
                HideSemanticConnector(page, semanticConnector);
                try
                {
                    // Keep endpoint arrows above the subject fill. Message
                    // containers are brought forward again after placement.
                    visualConnector.BringToFront();
                }
                catch (COMException)
                {
                    // The connector remains functional in protected documents
                    // that reject the z-order operation.
                }
            }
        }

        /// <summary>
        /// Replaces Visio's shortest-path result with an explicit orthogonal
        /// route through the corridor occupied by the corresponding message
        /// container.
        /// </summary>
        public static void RouteAlongMessageCorridors(
            Visio.IVPage page, LayoutDirection direction)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));

            List<Visio.Shape> semanticConnectors =
                new List<Visio.Shape>();
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (!IsVisualConnector(shape)
                    && VisioConnectorRebinder.IsStencilSidMessageConnector(
                        shape))
                {
                    semanticConnectors.Add(shape);
                }
            }

            foreach (Visio.Shape semanticConnector in semanticConnectors)
            {
                if (!VisioConnectorRebinder.TryGetConnectedShapes(
                    page, semanticConnector,
                    out Visio.Shape source, out Visio.Shape target))
                {
                    continue;
                }

                HideSemanticConnector(page, semanticConnector);
                Visio.Shape messageContainer = FindMessageContainer(
                    page, semanticConnector.ID);
                double[] points = BuildCorridorPoints(
                    source, target, messageContainer, direction);

                Visio.Shape previousVisual = FindVisualConnector(
                    page, semanticConnector.ID);
                Visio.Shape previousLeader = FindVisualLeader(
                    page, semanticConnector.ID);
                Visio.Shape replacement = CreateCorridorConnector(
                    page, semanticConnector, points);
                if (replacement == null) continue;

                MarkVisualConnector(
                    replacement, semanticConnector, source, target);
                TryBringToFront(replacement);

                Visio.Shape leader = CreateMessageLeader(
                    page, semanticConnector, messageContainer,
                    points, direction);
                if (leader != null)
                {
                    MarkVisualLeader(leader, semanticConnector);
                    TryBringToFront(leader);
                    if (previousLeader != null
                        && previousLeader.ID != leader.ID)
                    {
                        TryDelete(previousLeader);
                    }
                }

                BringMessageContainerToFront(page, messageContainer);
                if (previousVisual != null
                    && previousVisual.ID != replacement.ID)
                {
                    TryDelete(previousVisual);
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

        internal static bool IsVisualLeader(Visio.Shape shape)
        {
            return GetUserFlag(
                shape, Constants.UserCells.AutoArrangeSidVisualLeader);
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

        private static Visio.Shape FindVisualLeader(
            Visio.IVPage page, int semanticConnectorId)
        {
            string expectedId = semanticConnectorId.ToString(
                CultureInfo.InvariantCulture);
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (IsVisualLeader(shape)
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

        private static Visio.Shape CreateCorridorConnector(
            Visio.IVPage page, Visio.Shape semanticConnector,
            double[] points)
        {
            try
            {
                Array coordinateArray = points;
                Visio.Shape connector = page.DrawPolyline(
                    ref coordinateArray,
                    (short)Visio.VisDrawSplineFlags.visPolyline1D);
                CopyLineFormatting(semanticConnector, connector);
                // This shape is intentionally presentation-only. Gluing it or
                // marking it as a routable connector lets Visio discard the
                // explicit vertices and collapse the path onto a subject
                // border. The hidden stencil connector remains physically
                // glued and carries all semantic data.
                VisioRouting.TrySetCell(
                    connector, "ConFixedCode", 2, false);
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

        private static Visio.Shape CreateMessageLeader(
            Visio.IVPage page, Visio.Shape semanticConnector,
            Visio.Shape messageContainer, double[] points,
            LayoutDirection direction)
        {
            if (messageContainer == null || points == null
                || points.Length < 8)
            {
                return null;
            }

            try
            {
                double boxX = GetNumber(messageContainer, "PinX");
                double boxY = GetNumber(messageContainer, "PinY");
                double boxWidth = GetNumber(messageContainer, "Width");
                double boxHeight = GetNumber(messageContainer, "Height");
                string boxReference = "Sheet."
                    + messageContainer.ID.ToString(
                        CultureInfo.InvariantCulture) + "!";

                double beginX;
                double beginY;
                double endX;
                double endY;
                string beginXFormula;
                string beginYFormula;
                string endXFormula;
                string endYFormula;

                if (direction == LayoutDirection.TopDown)
                {
                    double corridorX = points[2];
                    double minimumY = Math.Min(points[3], points[5]);
                    double maximumY = Math.Max(points[3], points[5]);
                    beginX = boxX + (boxX < corridorX
                        ? boxWidth / 2d : -boxWidth / 2d);
                    beginY = boxY;
                    endX = corridorX;
                    endY = Clamp(boxY, minimumY, maximumY);
                    string corridor = FormatInches(corridorX);
                    beginXFormula = "GUARD(" + boxReference
                        + "PinX+IF(" + boxReference + "PinX<"
                        + corridor + ",0.5*" + boxReference
                        + "Width,-0.5*" + boxReference + "Width))";
                    beginYFormula = "GUARD(" + boxReference + "PinY)";
                    endXFormula = "GUARD(" + corridor + ")";
                    endYFormula = "GUARD(MAX(" + FormatInches(minimumY)
                        + ",MIN(" + FormatInches(maximumY) + ","
                        + boxReference + "PinY)))";
                }
                else
                {
                    double corridorY = points[3];
                    double minimumX = Math.Min(points[2], points[4]);
                    double maximumX = Math.Max(points[2], points[4]);
                    beginX = boxX;
                    beginY = boxY + (boxY < corridorY
                        ? boxHeight / 2d : -boxHeight / 2d);
                    endX = Clamp(boxX, minimumX, maximumX);
                    endY = corridorY;
                    string corridor = FormatInches(corridorY);
                    beginXFormula = "GUARD(" + boxReference + "PinX)";
                    beginYFormula = "GUARD(" + boxReference
                        + "PinY+IF(" + boxReference + "PinY<"
                        + corridor + ",0.5*" + boxReference
                        + "Height,-0.5*" + boxReference + "Height))";
                    endXFormula = "GUARD(MAX(" + FormatInches(minimumX)
                        + ",MIN(" + FormatInches(maximumX) + ","
                        + boxReference + "PinX)))";
                    endYFormula = "GUARD(" + corridor + ")";
                }

                Visio.Shape leader = page.DrawLine(
                    beginX, beginY, endX, endY);
                CopyLineFormatting(semanticConnector, leader);
                VisioRouting.TrySetCell(
                    leader, "ConFixedCode", 2, false);
                VisioRouting.TrySetCell(
                    leader, "ConLineJumpCode", 0, false);
                VisioRouting.TrySetCell(
                    leader, "BeginArrow", 0, false);
                VisioRouting.TrySetCell(
                    leader, "EndArrow", 0, false);
                leader.CellsU["BeginX"].FormulaForceU = beginXFormula;
                leader.CellsU["BeginY"].FormulaForceU = beginYFormula;
                leader.CellsU["EndX"].FormulaForceU = endXFormula;
                leader.CellsU["EndY"].FormulaForceU = endYFormula;
                return leader;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static string FormatInches(double value)
        {
            return value.ToString(CultureInfo.InvariantCulture) + " in";
        }

        private static double Clamp(double value, double minimum,
            double maximum)
        {
            return Math.Max(minimum, Math.Min(maximum, value));
        }

        private static double[] BuildCorridorPoints(
            Visio.Shape source, Visio.Shape target,
            Visio.Shape messageContainer, LayoutDirection direction)
        {
            double sourceX = GetNumber(source, "PinX");
            double sourceY = GetNumber(source, "PinY");
            double sourceWidth = GetNumber(source, "Width");
            double sourceHeight = GetNumber(source, "Height");
            double targetX = GetNumber(target, "PinX");
            double targetY = GetNumber(target, "PinY");
            double targetWidth = GetNumber(target, "Width");
            double targetHeight = GetNumber(target, "Height");
            bool selfLoop = source.ID == target.ID;

            if (direction == LayoutDirection.TopDown)
            {
                bool useLeft = !selfLoop
                    && sourceY < targetY - CoordinateTolerance;
                double sourceRelativeY = selfLoop ? 0.65 : 0.5;
                double targetRelativeY = selfLoop ? 0.35 : 0.5;

                double sourceEdgeX = sourceX
                    + (useLeft ? -sourceWidth / 2d : sourceWidth / 2d);
                double targetEdgeX = targetX
                    + (useLeft ? -targetWidth / 2d : targetWidth / 2d);
                double sourcePortY = sourceY
                    + (sourceRelativeY - 0.5) * sourceHeight;
                double targetPortY = targetY
                    + (targetRelativeY - 0.5) * targetHeight;
                double outsideX = useLeft
                    ? Math.Min(sourceEdgeX, targetEdgeX)
                        - CorridorClearance
                    : Math.Max(sourceEdgeX, targetEdgeX)
                        + CorridorClearance;
                double corridorX = messageContainer == null
                    ? outsideX
                    : GetNumber(messageContainer, "PinX");
                corridorX = useLeft
                    ? Math.Min(corridorX, outsideX)
                    : Math.Max(corridorX, outsideX);

                return new[]
                {
                    sourceEdgeX, sourcePortY,
                    corridorX, sourcePortY,
                    corridorX, targetPortY,
                    targetEdgeX, targetPortY
                };
            }

            bool useBottom = !selfLoop
                && sourceX > targetX + CoordinateTolerance;
            double sourceRelativeX = selfLoop ? 0.35 : 0.5;
            double targetRelativeX = selfLoop ? 0.65 : 0.5;

            double sourceEdgeY = sourceY
                + (useBottom ? -sourceHeight / 2d : sourceHeight / 2d);
            double targetEdgeY = targetY
                + (useBottom ? -targetHeight / 2d : targetHeight / 2d);
            double sourcePortX = sourceX
                + (sourceRelativeX - 0.5) * sourceWidth;
            double targetPortX = targetX
                + (targetRelativeX - 0.5) * targetWidth;
            double outsideY = useBottom
                ? Math.Min(sourceEdgeY, targetEdgeY) - CorridorClearance
                : Math.Max(sourceEdgeY, targetEdgeY) + CorridorClearance;
            double corridorY = messageContainer == null
                ? outsideY
                : GetNumber(messageContainer, "PinY");
            corridorY = useBottom
                ? Math.Min(corridorY, outsideY)
                : Math.Max(corridorY, outsideY);

            return new[]
            {
                sourcePortX, sourceEdgeY,
                sourcePortX, corridorY,
                targetPortX, corridorY,
                targetPortX, targetEdgeY
            };
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

        private static void MarkVisualLeader(Visio.Shape leader,
            Visio.Shape semanticConnector)
        {
            VisioShapeSheet.SetUserCell(leader,
                Constants.UserCells.AutoArrangeSidVisualLeader, 1);
            VisioShapeSheet.SetUserCell(leader,
                Constants.UserCells.AutoArrangeSidSemanticConnectorShapeId,
                semanticConnector.ID);
        }

        private static Visio.Shape FindMessageContainer(
            Visio.IVPage page, int semanticConnectorId)
        {
            const string correspondingShapeCell =
                "User.idOfCorrespondingShape";
            foreach (Visio.Shape shape in page.Shapes)
            {
                try
                {
                    if (shape.OneD != 0
                        || shape.CellExistsU[
                            correspondingShapeCell, 0] == 0)
                    {
                        continue;
                    }

                    int correspondingId = Convert.ToInt32(Math.Round(
                        shape.CellsU[correspondingShapeCell].Result[""]));
                    if (correspondingId == semanticConnectorId)
                        return shape;
                }
                catch (COMException)
                {
                    // Continue with the remaining page shapes.
                }
            }

            return null;
        }

        private static void BringMessageContainerToFront(
            Visio.IVPage page, Visio.Shape messageContainer)
        {
            if (messageContainer == null) return;

            try
            {
                messageContainer.BringToFront();
                Array memberIds = messageContainer.ContainerProperties
                    .GetListMembers();
                if (memberIds == null) return;

                foreach (object memberIdValue in memberIds)
                {
                    int memberId = Convert.ToInt32(memberIdValue);
                    if (memberId <= 0) continue;
                    page.Shapes.get_ItemFromID(memberId).BringToFront();
                }
            }
            catch (COMException)
            {
                // Routing stays valid on protected pages without z-ordering.
            }
        }

        private static void TryBringToFront(Visio.Shape shape)
        {
            try
            {
                shape?.BringToFront();
            }
            catch (COMException)
            {
                // Endpoint binding is independent of z-order.
            }
        }

        private static double GetNumber(
            Visio.Shape shape, string cellName)
        {
            return VisioShapeSheet.GetNumber(shape, cellName);
        }

        private static void HideSemanticConnector(
            Visio.IVPage page, Visio.Shape semanticConnector)
        {
            VisioShapeSheet.SetUserCell(semanticConnector,
                Constants.UserCells.AutoArrangeSidSemanticShadow, 1);

            // Older builds used line transparency as a fallback. The linked
            // MessageBox inherits that value, which also removed its border.
            // Restore the semantic line style and suppress only its geometry.
            TrySetFormulaForce(
                semanticConnector, "LineColorTrans", "0%");

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

                layer.CellsC[4].FormulaForceU = "0";
                layer.CellsC[5].FormulaForceU = "0";
                // Reassign every component of the grouped SmartShape. Keeping
                // its old subshape memberships can leave the internal link
                // geometries visible even though the group is on this layer.
                layer.Add(semanticConnector, (short)0);
            }
            catch (COMException)
            {
                // Geometry suppression below is independent of page layers.
            }

            HideGeometryRecursively(semanticConnector);
        }

        private static void HideGeometryRecursively(Visio.Shape shape)
        {
            // FormulaForceU is intentional: the legacy SID master protects
            // several geometry cells with GUARD. Geometry*.NoShow suppresses
            // both stroke and fill without changing the connector's semantic
            // properties, message-box formulas, or endpoint IDs.
            try
            {
                for (int index = 1; index <= shape.GeometryCount; index++)
                {
                    string noShowCell = "Geometry"
                        + index.ToString(CultureInfo.InvariantCulture)
                        + ".NoShow";
                    if (shape.CellExistsU[noShowCell, 0] != 0)
                        shape.CellsU[noShowCell].FormulaForceU = "TRUE";
                }
            }
            catch (COMException)
            {
                // Continue with child shapes; the hidden layer is independent.
            }
            try
            {
                foreach (Visio.Shape child in shape.Shapes)
                    HideGeometryRecursively(child);
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

        private static void TrySetFormulaForce(
            Visio.Shape shape, string cellName, string formula)
        {
            try
            {
                if (shape.CellExistsU[cellName, 0] != 0)
                    shape.CellsU[cellName].FormulaForceU = formula;
            }
            catch (COMException)
            {
                // Geometry*.NoShow and the hidden layer remain independent;
                // restoring one inherited style cell is non-fatal.
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
                // Prevent a failed cleanup from leaving duplicate geometry.
                HideGeometryRecursively(shape);
            }
        }
    }
}
