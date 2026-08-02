using System;
using System.Collections.Generic;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using LayoutDirection = ALPS_Visio_AddIn_rewrite.VisioHelper.GraphLayoutDirection;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    internal enum AlpsDiagramKind
    {
        Unknown,
        Sid,
        Sbd
    }

    /// <summary>
    /// Places ALPS/PASS nodes in deterministic process ranks without treating
    /// connector labels and SID message containers as graph nodes.
    /// </summary>
    internal static class VisioGraphAutoArranger
    {
        // ShapeSheet Result[""] values use Visio's internal inch units.
        private const double Margin = 0.8;
        private const double FeedbackCorridor = 0.65;
        // Connector labels in the PASS stencils are wider than their paths.
        // These gaps reserve a readable label corridor between ranks and
        // parallel alternatives instead of placing labels over sibling states.
        private const double FlowGap = 1.45;
        private const double SiblingGap = 1.25;
        private const double ConnectorLabelPadding = 0.45;
        private const double EstimatedCharacterWidth = 0.055;
        private const double MaximumConnectorLabelWidth = 3.1;
        private const double PortraitWidth = 8.2677;
        private const double PortraitHeight = 11.6929;
        private const double LandscapeWidth = PortraitHeight;
        private const double LandscapeHeight = PortraitWidth;

        public static bool TryArrange(Visio.IVPage page,
            LayoutDirection direction, AlpsDiagramKind diagramKind)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (diagramKind == AlpsDiagramKind.Unknown) return false;

            List<GraphNode> nodes = CollectNodes(page, diagramKind);
            if (nodes.Count == 0) return false;

            Dictionary<int, GraphNode> nodesById =
                nodes.ToDictionary(node => node.Shape.ID);
            CollectEdges(page, nodesById);
            AssignRanks(nodes, diagramKind);
            OrderRanks(nodes, direction);
            PlaceNodes(page, nodes, direction);
            return true;
        }

        private static List<GraphNode> CollectNodes(Visio.IVPage page,
            AlpsDiagramKind diagramKind)
        {
            List<GraphNode> nodes = new List<GraphNode>();
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.OneD != 0 || !IsDiagramNode(shape, diagramKind))
                    continue;

                nodes.Add(new GraphNode(
                    shape,
                    GetStableKey(shape),
                    VisioShapeSheet.GetNumber(shape, "Width"),
                    VisioShapeSheet.GetNumber(shape, "Height")));
            }

            return nodes
                .OrderBy(node => node.StableKey, StringComparer.OrdinalIgnoreCase)
                .ThenBy(node => node.Shape.ID)
                .ToList();
        }

        private static void CollectEdges(Visio.IVPage page,
            IDictionary<int, GraphNode> nodesById)
        {
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (VisioSidMessageConnectorRenderer.IsSemanticShadow(shape)
                    || !VisioConnectorRebinder.HasConnectorEndpoints(shape)
                    || !VisioConnectorRebinder.TryGetConnectedShapes(
                        page, shape, out Visio.Shape sourceShape,
                        out Visio.Shape targetShape)
                    || !nodesById.TryGetValue(
                        sourceShape.ID, out GraphNode source)
                    || !nodesById.TryGetValue(
                        targetShape.ID, out GraphNode target))
                {
                    continue;
                }

                GraphEdge edge = new GraphEdge(shape, source, target);
                source.Outgoing.Add(edge);
                target.Incoming.Add(edge);
            }
        }

        private static void AssignRanks(
            IList<GraphNode> nodes, AlpsDiagramKind diagramKind)
        {
            foreach (GraphNode node in nodes)
                node.Rank = -1;

            List<GraphNode> roots = nodes
                .Where(node => IsStartNode(node.Shape, diagramKind))
                .ToList();
            if (roots.Count == 0)
            {
                roots = nodes
                    .Where(node => node.Incoming.Count == 0)
                    .ToList();
            }

            if (roots.Count == 0)
                roots.Add(nodes.OrderBy(node => node.StableKey,
                    StringComparer.OrdinalIgnoreCase).ThenBy(
                    node => node.Shape.ID).First());

            // Breadth-first ranks keep cyclic edges as feedback instead of
            // collapsing a complete behavior cycle into a single rank.
            AssignReachableRanks(roots, 0);

            while (nodes.Any(node => node.Rank < 0))
            {
                int nextRank = nodes.Where(node => node.Rank >= 0)
                    .Select(node => node.Rank).DefaultIfEmpty(-1).Max() + 1;
                List<GraphNode> remainingRoots = nodes
                    .Where(node => node.Rank < 0)
                    .Where(node => node.Incoming.All(
                        edge => edge.Source.Rank >= 0))
                    .ToList();
                if (remainingRoots.Count == 0)
                {
                    remainingRoots.Add(nodes
                        .Where(node => node.Rank < 0)
                        .OrderBy(node => node.StableKey,
                            StringComparer.OrdinalIgnoreCase)
                        .ThenBy(node => node.Shape.ID)
                        .First());
                }

                AssignReachableRanks(remainingRoots, nextRank);
            }
        }

        private static void AssignReachableRanks(
            IEnumerable<GraphNode> roots, int initialRank)
        {
            Queue<GraphNode> queue = new Queue<GraphNode>();
            foreach (GraphNode root in roots
                .OrderBy(node => node.StableKey,
                    StringComparer.OrdinalIgnoreCase)
                .ThenBy(node => node.Shape.ID))
            {
                if (root.Rank >= 0) continue;
                root.Rank = initialRank;
                queue.Enqueue(root);
            }

            while (queue.Count > 0)
            {
                GraphNode current = queue.Dequeue();
                foreach (GraphNode target in current.Outgoing
                    .Select(edge => edge.Target)
                    .Distinct()
                    .OrderBy(node => node.StableKey,
                        StringComparer.OrdinalIgnoreCase)
                    .ThenBy(node => node.Shape.ID))
                {
                    if (target.Rank >= 0) continue;
                    target.Rank = current.Rank + 1;
                    queue.Enqueue(target);
                }
            }
        }

        private static void OrderRanks(
            IList<GraphNode> nodes, LayoutDirection direction)
        {
            List<IGrouping<int, GraphNode>> rankGroups = nodes
                .GroupBy(node => node.Rank)
                .OrderBy(group => group.Key)
                .ToList();

            foreach (IGrouping<int, GraphNode> rankGroup in rankGroups)
            {
                IEnumerable<GraphNode> ordered = direction
                    == LayoutDirection.TopDown
                    ? rankGroup.OrderBy(node => GetX(node.Shape))
                    : rankGroup.OrderByDescending(node => GetY(node.Shape));
                SetOrder(ordered);
            }

            for (int sweep = 0; sweep < 3; sweep++)
            {
                // Alternating barycentric sweeps reduce crossings at branches
                // and joins while preserving a deterministic tie order.
                foreach (IGrouping<int, GraphNode> rankGroup in rankGroups)
                {
                    SetOrder(rankGroup
                        .OrderBy(node => GetBarycenter(
                            node.Incoming
                                .Where(edge =>
                                    edge.Source.Rank < node.Rank)
                                .Select(edge => edge.Source),
                            node.Order))
                        .ThenBy(node => node.Order)
                        .ThenBy(node => node.StableKey,
                            StringComparer.OrdinalIgnoreCase));
                }

                foreach (IGrouping<int, GraphNode> rankGroup
                    in rankGroups.AsEnumerable().Reverse())
                {
                    SetOrder(rankGroup
                        .OrderBy(node => GetBarycenter(
                            node.Outgoing
                                .Where(edge =>
                                    edge.Target.Rank > node.Rank)
                                .Select(edge => edge.Target),
                            node.Order))
                        .ThenBy(node => node.Order)
                        .ThenBy(node => node.StableKey,
                            StringComparer.OrdinalIgnoreCase));
                }
            }
        }

        private static void SetOrder(IEnumerable<GraphNode> orderedNodes)
        {
            int order = 0;
            foreach (GraphNode node in orderedNodes)
                node.Order = order++;
        }

        private static double GetBarycenter(
            IEnumerable<GraphNode> adjacentNodes, int fallback)
        {
            List<GraphNode> adjacent = adjacentNodes.Distinct().ToList();
            return adjacent.Count == 0
                ? fallback
                : adjacent.Average(node => node.Order);
        }

        private static void PlaceNodes(Visio.IVPage page,
            IList<GraphNode> nodes, LayoutDirection direction)
        {
            List<RankBand> bands = nodes
                .GroupBy(node => node.Rank)
                .OrderBy(group => group.Key)
                .Select(group => new RankBand(
                    group.OrderBy(node => node.Order).ToList(), direction))
                .ToList();
            ConfigureFlowGaps(bands, direction);

            double contentFlow = bands.Sum(band => band.FlowSize)
                + bands.Sum(band => band.GapAfter);
            double contentCross = bands.Max(band => band.CrossSize);
            double reservedCross = contentCross
                + 2d * (Margin + FeedbackCorridor);

            double pageWidth;
            double pageHeight;
            if (direction == LayoutDirection.TopDown)
            {
                pageWidth = Math.Max(PortraitWidth, reservedCross);
                pageHeight = Math.Max(PortraitHeight,
                    contentFlow + 2d * Margin);
            }
            else
            {
                pageWidth = Math.Max(LandscapeWidth,
                    contentFlow + 2d * Margin);
                pageHeight = Math.Max(LandscapeHeight, reservedCross);
            }

            VisioShapeSheet.SetNumber(page.PageSheet, "PageWidth", pageWidth);
            VisioShapeSheet.SetNumber(page.PageSheet, "PageHeight", pageHeight);

            if (direction == LayoutDirection.TopDown)
                PlaceTopDown(bands, pageWidth, pageHeight, contentFlow);
            else
                PlaceLeftRight(bands, pageWidth, pageHeight, contentFlow);
        }

        private static void ConfigureFlowGaps(
            IList<RankBand> bands, LayoutDirection direction)
        {
            for (int index = 0; index < bands.Count - 1; index++)
            {
                RankBand current = bands[index];
                RankBand next = bands[index + 1];
                current.GapAfter = direction == LayoutDirection.LeftRight
                    ? GetLeftRightFlowGap(current, next)
                    : FlowGap;
            }
        }

        private static double GetLeftRightFlowGap(
            RankBand current, RankBand next)
        {
            ISet<int> nextNodeIds = new HashSet<int>(
                next.Nodes.Select(node => node.Shape.ID));
            double widestLabel = current.Nodes
                .SelectMany(node => node.Outgoing)
                .Where(edge => nextNodeIds.Contains(edge.Target.Shape.ID))
                .Select(edge => EstimateConnectorLabelWidth(edge.Connector))
                .DefaultIfEmpty(0d)
                .Max();

            return Math.Max(
                FlowGap, widestLabel + ConnectorLabelPadding);
        }

        private static double EstimateConnectorLabelWidth(
            Visio.Shape connector)
        {
            int longestLine = GetLongestTextLine(connector);
            return Math.Min(
                MaximumConnectorLabelWidth,
                longestLine * EstimatedCharacterWidth);
        }

        private static int GetLongestTextLine(Visio.Shape shape)
        {
            int longestLine = 0;
            try
            {
                string text = shape.Text;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    longestLine = text
                        .Replace("\r\n", "\n")
                        .Replace('\r', '\n')
                        .Split('\n')
                        .Select(line => line.Trim().Length)
                        .DefaultIfEmpty(0)
                        .Max();
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Some master subshapes do not expose a text range.
            }

            try
            {
                foreach (Visio.Shape child in shape.Shapes)
                {
                    longestLine = Math.Max(
                        longestLine, GetLongestTextLine(child));
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Atomic shapes do not have child shapes.
            }

            return longestLine;
        }

        private static void PlaceTopDown(IEnumerable<RankBand> bands,
            double pageWidth, double pageHeight, double contentFlow)
        {
            double top = (pageHeight + contentFlow) / 2d;
            foreach (RankBand band in bands)
            {
                double centerY = top - band.FlowSize / 2d;
                double left = (pageWidth - band.CrossSize) / 2d;
                foreach (GraphNode node in band.Nodes)
                {
                    double centerX = left + node.Width / 2d;
                    SetPosition(node.Shape, centerX, centerY);
                    left += node.Width + SiblingGap;
                }

                top -= band.FlowSize + band.GapAfter;
            }
        }

        private static void PlaceLeftRight(IEnumerable<RankBand> bands,
            double pageWidth, double pageHeight, double contentFlow)
        {
            double left = (pageWidth - contentFlow) / 2d;
            foreach (RankBand band in bands)
            {
                double centerX = left + band.FlowSize / 2d;
                double top = (pageHeight + band.CrossSize) / 2d;
                foreach (GraphNode node in band.Nodes)
                {
                    double centerY = top - node.Height / 2d;
                    SetPosition(node.Shape, centerX, centerY);
                    top -= node.Height + SiblingGap;
                }

                left += band.FlowSize + band.GapAfter;
            }
        }

        private static void SetPosition(
            Visio.Shape shape, double x, double y)
        {
            VisioShapeSheet.SetNumber(shape, "PinX", x);
            VisioShapeSheet.SetNumber(shape, "PinY", y);
        }

        private static bool IsDiagramNode(
            Visio.Shape shape, AlpsDiagramKind diagramKind)
        {
            if (diagramKind == AlpsDiagramKind.Sbd)
            {
                return HasCategory(shape, Constants.ShapeCategories.SBDState)
                    || HasMaster(shape,
                        Constants.SBDMasters.DoState,
                        Constants.SBDMasters.ReceiveState,
                        Constants.SBDMasters.SendState,
                        Constants.SBDMasters.GenericReturnToOriginReference);
            }

            return HasCategory(shape, Constants.ShapeCategories.SIDSubject)
                || HasCategory(
                    shape, Constants.ShapeCategories.SIDSubjectWithSBD)
                || HasMaster(shape,
                    Constants.SIDMasters.StandardActor,
                    Constants.SIDMasters.InterfaceActor,
                    Constants.SIDMasters.StandAloneMacro,
                    Constants.SIDMasters.ActorExtension,
                    Constants.SIDMasters.SystemInterfaceSubject,
                    Constants.SIDMasters.SubjectGroup);
        }

        private static bool IsStartNode(
            Visio.Shape shape, AlpsDiagramKind diagramKind)
        {
            return GetBooleanProperty(shape,
                diagramKind == AlpsDiagramKind.Sbd
                    ? Constants.Properties.State.Start
                    : Constants.Properties.Subject.Start);
        }

        private static bool GetBooleanProperty(
            Visio.Shape shape, string propertyName)
        {
            string cellName = "Prop." + propertyName;
            try
            {
                return shape.CellExistsU[cellName, 0] != 0
                    && Math.Abs(shape.CellsU[cellName].Result[""]) > 0.5;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private static string GetStableKey(Visio.Shape shape)
        {
            string idCell = "Prop." + Constants.Properties.ID;
            try
            {
                if (shape.CellExistsU[idCell, 0] != 0)
                {
                    string id = shape.CellsU[idCell].ResultStr[""];
                    if (!string.IsNullOrWhiteSpace(id)) return id;
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // A shape name remains a deterministic fallback.
            }

            return shape.NameU ?? shape.ID.ToString();
        }

        private static bool HasCategory(
            Visio.Shape shape, string category)
        {
            try
            {
                return shape.HasCategory(category);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private static bool HasMaster(
            Visio.Shape shape, params string[] masterNames)
        {
            try
            {
                Visio.Master master = shape.Master;
                if (master == null) return false;
                return masterNames.Any(name => string.Equals(
                    name, master.NameU, StringComparison.OrdinalIgnoreCase));
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        private static double GetX(Visio.Shape shape)
        {
            return VisioShapeSheet.GetNumber(shape, "PinX");
        }

        private static double GetY(Visio.Shape shape)
        {
            return VisioShapeSheet.GetNumber(shape, "PinY");
        }

        private sealed class GraphNode
        {
            public GraphNode(Visio.Shape shape, string stableKey,
                double width, double height)
            {
                Shape = shape;
                StableKey = stableKey;
                Width = width;
                Height = height;
                Incoming = new List<GraphEdge>();
                Outgoing = new List<GraphEdge>();
            }

            public Visio.Shape Shape { get; private set; }
            public string StableKey { get; private set; }
            public double Width { get; private set; }
            public double Height { get; private set; }
            public IList<GraphEdge> Incoming { get; private set; }
            public IList<GraphEdge> Outgoing { get; private set; }
            public int Rank { get; set; }
            public int Order { get; set; }
        }

        private sealed class GraphEdge
        {
            public GraphEdge(
                Visio.Shape connector, GraphNode source, GraphNode target)
            {
                Connector = connector;
                Source = source;
                Target = target;
            }

            public Visio.Shape Connector { get; private set; }
            public GraphNode Source { get; private set; }
            public GraphNode Target { get; private set; }
        }

        private sealed class RankBand
        {
            public RankBand(
                IList<GraphNode> nodes, LayoutDirection direction)
            {
                Nodes = nodes;
                FlowSize = direction == LayoutDirection.TopDown
                    ? nodes.Max(node => node.Height)
                    : nodes.Max(node => node.Width);
                CrossSize = direction == LayoutDirection.TopDown
                    ? nodes.Sum(node => node.Width)
                        + Math.Max(0, nodes.Count - 1) * SiblingGap
                    : nodes.Sum(node => node.Height)
                        + Math.Max(0, nodes.Count - 1) * SiblingGap;
                GapAfter = 0d;
            }

            public IList<GraphNode> Nodes { get; private set; }
            public double FlowSize { get; private set; }
            public double CrossSize { get; private set; }
            public double GapAfter { get; set; }
        }
    }
}
