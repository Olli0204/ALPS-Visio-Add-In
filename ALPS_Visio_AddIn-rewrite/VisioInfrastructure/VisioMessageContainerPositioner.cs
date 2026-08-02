using System;
using System.Collections.Generic;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using LayoutDirection = ALPS_Visio_AddIn_rewrite.VisioHelper.GraphLayoutDirection;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Keeps SID message lists close to their communication channel after
    /// subject nodes have moved.
    /// </summary>
    internal static class VisioMessageContainerPositioner
    {
        private const string MessageCategory = "alpsMessage";
        private const double ConnectorClearance = 0.45;
        private const double ContainerGap = 0.25;
        private const double PageClearance = 0.35;
        private const double ContainmentTolerance = 0.05;

        public static void Reposition(
            Visio.IVPage page, LayoutDirection direction)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));

            Dictionary<int, Visio.Shape> connectors =
                new Dictionary<int, Visio.Shape>();
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (VisioConnectorRebinder.HasConnectorEndpoints(shape))
                    connectors[shape.ID] = shape;
            }
            double pageWidth =
                VisioShapeSheet.GetNumber(page.PageSheet, "PageWidth");
            double pageHeight =
                VisioShapeSheet.GetNumber(page.PageSheet, "PageHeight");
            List<ContainerPlacement> placements =
                new List<ContainerPlacement>();

            foreach (Visio.Shape container in page.Shapes)
            {
                if (container.OneD != 0 || !IsMessageContainer(container)
                    || !TryGetCorrespondingShapeId(
                        container, out int connectorId)
                    || !connectors.TryGetValue(
                        connectorId, out Visio.Shape connector)
                    || !VisioConnectorRebinder.TryGetConnectedShapes(
                        page, connector, out Visio.Shape source,
                        out Visio.Shape target))
                {
                    continue;
                }

                placements.Add(new ContainerPlacement(
                    container, connector, source, target));
            }

            foreach (IGrouping<string, ContainerPlacement> placementGroup
                in placements.GroupBy(placement => placement.PairKey))
            {
                List<ContainerPlacement> orderedPlacements = placementGroup
                    .OrderBy(placement => placement.Connector.ID)
                    .ToList();
                for (int index = 0;
                    index < orderedPlacements.Count; index++)
                {
                    ContainerPlacement placement = orderedPlacements[index];
                    double sourceX = GetNumber(placement.Source, "PinX");
                    double sourceY = GetNumber(placement.Source, "PinY");
                    double targetX = GetNumber(placement.Target, "PinX");
                    double targetY = GetNumber(placement.Target, "PinY");
                    double width = GetNumber(placement.Container, "Width");
                    double height = GetNumber(placement.Container, "Height");
                    double x = (sourceX + targetX) / 2d;
                    double y = (sourceY + targetY) / 2d;
                    double laneIndex =
                        index - (orderedPlacements.Count - 1d) / 2d;

                    if (direction == LayoutDirection.TopDown)
                    {
                        bool feedback = sourceY < targetY;
                        double endpointWidth = Math.Max(
                            GetNumber(placement.Source, "Width"),
                            GetNumber(placement.Target, "Width"));
                        x += (feedback ? -1d : 1d)
                            * (endpointWidth / 2d + width / 2d
                                + ConnectorClearance);
                        y += laneIndex * (height + ContainerGap);
                    }
                    else
                    {
                        bool feedback = sourceX > targetX;
                        double endpointHeight = Math.Max(
                            GetNumber(placement.Source, "Height"),
                            GetNumber(placement.Target, "Height"));
                        y += (feedback ? -1d : 1d)
                            * (endpointHeight / 2d + height / 2d
                                + ConnectorClearance);
                        x += laneIndex * (width + ContainerGap);
                    }

                    x = Clamp(x, width / 2d + PageClearance,
                        pageWidth - width / 2d - PageClearance);
                    y = Clamp(y, height / 2d + PageClearance,
                        pageHeight - height / 2d - PageClearance);
                    MoveContainerWithListMembers(
                        page, placement.Container, x, y);
                }
            }
        }

        private static void MoveContainerWithListMembers(
            Visio.IVPage page, Visio.Shape container, double x, double y)
        {
            List<Visio.Shape> members =
                CaptureListMembers(page, container);
            if (members.Count == 0)
            {
                // Auto-Arrange versions before this fix could leave the
                // messages visually inside the box but no longer registered
                // as list members. Recover those diagrams on the next run.
                members = CaptureContainedMessages(page, container);
            }

            RemoveListMembers(container, members);
            VisioShapeSheet.SetNumber(container, "PinX", x);
            VisioShapeSheet.SetNumber(container, "PinY", y);
            try
            {
                // Native SID connectors must stay above the subject fills so
                // their arrowheads remain visible. Keep the message container
                // above those connectors so no route crosses its label area.
                container.BringToFront();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Positioning and list recovery still work on protected pages.
            }

            int listPosition = 1;
            foreach (Visio.Shape member in members)
            {
                // Inserting the shapes again lets Visio move and format them
                // as true list members. Direct PinX/PinY changes detach them
                // and make the Message master show its red warning outline.
                container.ContainerProperties.InsertListMember(
                    member, listPosition++);
                member.BringToFront();
            }
        }

        private static List<Visio.Shape> CaptureListMembers(
            Visio.IVPage page, Visio.Shape container)
        {
            List<Visio.Shape> members = new List<Visio.Shape>();
            try
            {
                Array memberIds =
                    container.ContainerProperties.GetListMembers();
                if (memberIds == null) return members;

                foreach (object memberIdValue in memberIds)
                {
                    int memberId = Convert.ToInt32(memberIdValue);
                    if (memberId <= 0 || memberId == container.ID) continue;

                    Visio.Shape member =
                        page.Shapes.get_ItemFromID(memberId);
                    members.Add(member);
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // A malformed or non-list container can still be positioned.
            }

            return members;
        }

        private static List<Visio.Shape> CaptureContainedMessages(
            Visio.IVPage page, Visio.Shape container)
        {
            double centerX = GetNumber(container, "PinX");
            double centerY = GetNumber(container, "PinY");
            double halfWidth = GetNumber(container, "Width") / 2d
                + ContainmentTolerance;
            double halfHeight = GetNumber(container, "Height") / 2d
                + ContainmentTolerance;
            List<Visio.Shape> messages = new List<Visio.Shape>();

            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.OneD == 0
                    && shape.ID != container.ID
                    && IsMessage(shape)
                    && Math.Abs(GetNumber(shape, "PinX") - centerX)
                        <= halfWidth
                    && Math.Abs(GetNumber(shape, "PinY") - centerY)
                        <= halfHeight)
                {
                    messages.Add(shape);
                }
            }

            return messages
                .OrderByDescending(shape => GetNumber(shape, "PinY"))
                .ThenBy(shape => GetNumber(shape, "PinX"))
                .ThenBy(shape => shape.ID)
                .ToList();
        }

        private static void RemoveListMembers(
            Visio.Shape container, IEnumerable<Visio.Shape> members)
        {
            foreach (Visio.Shape member in members)
            {
                try
                {
                    container.ContainerProperties.RemoveMember(member);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // A recovered visual member may already be detached.
                }
            }
        }

        private static bool IsMessage(Visio.Shape shape)
        {
            try
            {
                if (shape.HasCategory(MessageCategory)) return true;

                Visio.Master master = shape.Master;
                return master != null && string.Equals(
                    master.NameU, Constants.SIDMasters.Message,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private static bool IsMessageContainer(Visio.Shape shape)
        {
            try
            {
                if (shape.HasCategory(
                    Constants.ShapeCategories.SIDMessageConnectorBox))
                {
                    return true;
                }

                Visio.Master master = shape.Master;
                return master != null && string.Equals(
                    master.NameU, Constants.SIDMasters.MessageBox,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private static bool TryGetCorrespondingShapeId(
            Visio.Shape shape, out int shapeId)
        {
            shapeId = 0;
            const string cellName = "User.idOfCorrespondingShape";
            try
            {
                if (shape.CellExistsU[cellName, 0] == 0) return false;
                shapeId = Convert.ToInt32(
                    Math.Round(shape.CellsU[cellName].Result[""]));
                return shapeId > 0;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private static double GetNumber(
            Visio.Shape shape, string cellName)
        {
            return VisioShapeSheet.GetNumber(shape, cellName);
        }

        private static double Clamp(double value, double minimum,
            double maximum)
        {
            if (maximum < minimum) return (minimum + maximum) / 2d;
            return Math.Max(minimum, Math.Min(maximum, value));
        }

        private sealed class ContainerPlacement
        {
            public ContainerPlacement(Visio.Shape container,
                Visio.Shape connector, Visio.Shape source, Visio.Shape target)
            {
                Container = container;
                Connector = connector;
                Source = source;
                Target = target;
                PairKey = source.ID.ToString() + ">" + target.ID.ToString();
            }

            public Visio.Shape Container { get; private set; }
            public Visio.Shape Connector { get; private set; }
            public Visio.Shape Source { get; private set; }
            public Visio.Shape Target { get; private set; }
            public string PairKey { get; private set; }
        }

    }
}
