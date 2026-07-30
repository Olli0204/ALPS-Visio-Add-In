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
        private const double ConnectorClearance = 0.45;
        private const double ContainerGap = 0.25;
        private const double PageClearance = 0.35;

        public static void Reposition(
            Visio.IVPage page, LayoutDirection direction)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));

            Dictionary<int, Visio.Shape> connectors =
                new Dictionary<int, Visio.Shape>();
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.OneD != 0)
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
                        connector, out Visio.Shape source,
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
            double deltaX = x - GetNumber(container, "PinX");
            double deltaY = y - GetNumber(container, "PinY");
            List<MemberPosition> members =
                CaptureListMemberPositions(page, container);

            VisioShapeSheet.SetNumber(container, "PinX", x);
            VisioShapeSheet.SetNumber(container, "PinY", y);

            // List members are independent page shapes. Moving a container
            // through ShapeSheet cells does not reliably move them with it.
            // Explicit target positions also avoid applying the delta twice
            // in Visio versions that already move members automatically.
            foreach (MemberPosition member in members)
            {
                VisioShapeSheet.SetNumber(
                    member.Shape, "PinX", member.X + deltaX);
                VisioShapeSheet.SetNumber(
                    member.Shape, "PinY", member.Y + deltaY);
            }
        }

        private static List<MemberPosition> CaptureListMemberPositions(
            Visio.IVPage page, Visio.Shape container)
        {
            List<MemberPosition> members = new List<MemberPosition>();
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
                    members.Add(new MemberPosition(
                        member, GetNumber(member, "PinX"),
                        GetNumber(member, "PinY")));
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // A malformed or non-list container can still be positioned.
            }

            return members;
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

        private sealed class MemberPosition
        {
            public MemberPosition(
                Visio.Shape shape, double x, double y)
            {
                Shape = shape;
                X = x;
                Y = y;
            }

            public Visio.Shape Shape { get; private set; }
            public double X { get; private set; }
            public double Y { get; private set; }
        }
    }
}
