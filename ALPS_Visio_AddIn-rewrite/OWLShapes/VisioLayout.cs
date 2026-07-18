using alps.net.api.ALPS;
using alps.net.api.StandardPASS;
using System.Collections.Generic;
using System.Linq;
using System;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    /// <summary>
    /// Supplies deterministic graph-based fallback geometry for model elements
    /// that do not contain ALPS 2D visualisation data.
    /// </summary>
    internal static class VisioLayout
    {
        private const double Margin = 0.12;
        private const double SubjectHorizontalMargin = 0.26;
        private const double StateHorizontalMargin = 0.12;
        private const double SubjectWidth = 0.22;
        private const double SubjectHeight = 0.16;
        private const double StateWidth = 0.14;
        private const double StateHeight = 0.06;
        private const double StateConnectionPitch = 0.015;
        private const double StateRowGap = 0.10;
        private const double RegularPortMargin = 0.25;
        private const double FeedbackPortPosition = 0.10;
        private static readonly IDictionary<IPASSProcessModelElement, LayoutBounds> GeneratedBounds =
            new Dictionary<IPASSProcessModelElement, LayoutBounds>();
        private static readonly IDictionary<string, TransitionPorts> PortsByTransitionId =
            new Dictionary<string, TransitionPorts>();

        public static bool ArrangeModelLayer(IEnumerable<IPASSProcessModelElement> elements)
        {
            GeneratedBounds.Clear();
            List<ISubject> subjects = elements.OfType<ISubject>().ToList();
            Dictionary<ISubject, int> ranks = DetermineSubjectRanks(subjects);
            return ArrangeSubjects(subjects, ranks);
        }

        public static bool ArrangeBehavior(IEnumerable<IBehaviorDescribingComponent> components)
        {
            List<IState> states = components.OfType<IState>().ToList();
            List<ITransition> transitions = components.OfType<ITransition>().ToList();
            PortsByTransitionId.Clear();
            Dictionary<IState, int> ranks = DetermineStateRanks(states, transitions);
            Dictionary<IState, int> verticalOrder = DetermineStateVerticalOrder(states, transitions, ranks);
            PrepareTransitionPorts(states, transitions, ranks, verticalOrder);
            bool fallbackApplied = ArrangeStates(states, transitions, ranks, verticalOrder);
            MarkFallbackTransitions(transitions);
            return fallbackApplied;
        }

        internal static void GetTransitionPorts(ITransition transition, out double sourceY, out double targetY,
            out bool isFeedback, out bool useFallbackRouting)
        {
            sourceY = 0.5;
            targetY = 0.5;
            isFeedback = false;
            useFallbackRouting = false;
            if (transition == null || string.IsNullOrEmpty(transition.getModelComponentID())) return;

            if (PortsByTransitionId.TryGetValue(transition.getModelComponentID(), out TransitionPorts ports)
                && ports.UseFallbackRouting)
            {
                sourceY = ports.SourceY;
                targetY = ports.TargetY;
                isFeedback = ports.IsFeedback;
                useFallbackRouting = true;
            }
        }

        public static void PrepareOrArrange(IVisioExportableWithShape exportable, int index, bool subjectDiagram)
        {
            IPASSProcessModelElement element = exportable as IPASSProcessModelElement;
            if (exportable == null || element == null || HasGeneratedBounds(element)) return;
            if (PrepareUsableBounds(exportable, element)) return;

            int columns = subjectDiagram ? 4 : 3;
            int column = index % columns;
            int row = index / columns;
            double x = Margin + column * ((1.0 - 2 * Margin) / Math.Max(1, columns - 1));
            double y = Math.Max(Margin, 1.0 - Margin - row * 0.20);
            SetBounds(element, x, y, subjectDiagram ? StateWidth : SubjectWidth,
                subjectDiagram ? StateHeight : SubjectHeight);
        }

        private static bool ArrangeSubjects(IEnumerable<ISubject> subjects, IDictionary<ISubject, int> ranks)
        {
            bool fallbackApplied = false;
            int maxRank = ranks.Count == 0 ? 0 : ranks.Values.Max();
            double width = GetNodeWidth(SubjectWidth, maxRank, SubjectHorizontalMargin);
            foreach (IGrouping<int, ISubject> rankGroup in subjects.GroupBy(subject => ranks[subject]).OrderBy(group => group.Key))
            {
                List<ISubject> missingSubjects = rankGroup
                    .Where(subject => NeedsFallbackBounds(subject as IVisioExportableWithShape))
                    .OrderByDescending(GetMessageExchangeCount)
                    .ThenBy(subject => subject.getModelComponentID())
                    .ToList();

                for (int row = 0; row < missingSubjects.Count; row++)
                {
                    double x = GetRankPosition(rankGroup.Key, maxRank, SubjectHorizontalMargin);
                    double y = GetRowPosition(row, missingSubjects.Count);
                    SetBounds((IPASSProcessModelElement)missingSubjects[row], x, y, width,
                        GetNodeHeight(SubjectHeight, missingSubjects.Count));
                    fallbackApplied = true;
                }
            }

            return fallbackApplied;
        }

        private static bool ArrangeStates(IEnumerable<IState> states, IEnumerable<ITransition> transitions,
            IDictionary<IState, int> ranks, IDictionary<IState, int> verticalOrder)
        {
            bool fallbackApplied = false;
            List<ITransition> transitionList = transitions.ToList();
            int maxRank = ranks.Count == 0 ? 0 : ranks.Values.Max();
            double width = GetNodeWidth(StateWidth, maxRank, StateHorizontalMargin);
            foreach (IGrouping<int, IState> rankGroup in states.GroupBy(state => ranks[state]).OrderBy(group => group.Key))
            {
                List<IState> missingStates = rankGroup
                    .Where(state => NeedsFallbackBounds(state as IVisioExportableWithShape))
                    .OrderBy(state => verticalOrder[state])
                    .ToList();

                List<double> heights = missingStates
                    .Select(state => GetRequiredStateHeight(state, transitionList))
                    .ToList();
                double rowGap = StateRowGap;
                ScaleRowsToAvailableHeight(heights, ref rowGap);
                double usedHeight = heights.Sum() + Math.Max(0, heights.Count - 1) * rowGap;
                double top = 0.5 + usedHeight / 2.0;

                for (int row = 0; row < missingStates.Count; row++)
                {
                    double x = GetRankPosition(rankGroup.Key, maxRank, StateHorizontalMargin);
                    double y = top - heights[row] / 2.0;
                    SetBounds((IPASSProcessModelElement)missingStates[row], x, y, width, heights[row]);
                    top -= heights[row] + rowGap;
                    fallbackApplied = true;
                }
            }

            return fallbackApplied;
        }

        private static Dictionary<ISubject, int> DetermineSubjectRanks(IEnumerable<ISubject> subjects)
        {
            List<ISubject> subjectList = subjects.ToList();
            Dictionary<ISubject, int> ranks = subjectList.ToDictionary(subject => subject, subject => 0);
            Dictionary<ISubject, int> incomingCounts = subjectList.ToDictionary(subject => subject, subject => 0);

            foreach (ISubject subject in subjectList)
                foreach (IMessageExchange exchange in subject.getOutgoingMessageExchanges().Values)
                    if (exchange.getReceiver() != null && incomingCounts.ContainsKey(exchange.getReceiver()))
                        incomingCounts[exchange.getReceiver()]++;

            Queue<ISubject> queue = new Queue<ISubject>(incomingCounts
                .Where(pair => pair.Value == 0)
                .Select(pair => pair.Key)
                .OrderBy(candidate => candidate.getModelComponentID()));
            HashSet<ISubject> processed = new HashSet<ISubject>();
            while (processed.Count < subjectList.Count)
            {
                // A bidirectional exchange produces a cycle and has no natural
                // root. Pick a stable root so connected subjects receive
                // separate columns instead of being stacked on top of each other.
                if (queue.Count == 0)
                {
                    ISubject cycleRoot = subjectList
                        .Where(candidate => !processed.Contains(candidate))
                        .OrderByDescending(GetMessageExchangeCount)
                        .ThenBy(candidate => candidate.getModelComponentID())
                        .First();
                    queue.Enqueue(cycleRoot);
                }

                ISubject subject = queue.Dequeue();
                if (!processed.Add(subject)) continue;
                foreach (IMessageExchange exchange in subject.getOutgoingMessageExchanges().Values)
                {
                    ISubject receiver = exchange.getReceiver();
                    if (receiver == null || processed.Contains(receiver) || !incomingCounts.ContainsKey(receiver)) continue;

                    ranks[receiver] = Math.Max(ranks[receiver], ranks[subject] + 1);
                    incomingCounts[receiver]--;
                    if (incomingCounts[receiver] == 0)
                    {
                        queue.Enqueue(receiver);
                    }
                }
            }

            return ranks;
        }

        private static void PrepareTransitionPorts(IEnumerable<IState> states, IEnumerable<ITransition> transitions,
            IDictionary<IState, int> ranks, IDictionary<IState, int> verticalOrder)
        {
            List<IState> stateList = states.ToList();
            Dictionary<IState, List<ITransition>> outgoing =
                stateList.ToDictionary(candidate => candidate, candidate => new List<ITransition>());
            Dictionary<IState, List<ITransition>> incoming =
                stateList.ToDictionary(candidate => candidate, candidate => new List<ITransition>());

            foreach (ITransition transition in transitions)
            {
                IState source = transition.getSourceState();
                IState target = transition.getTargetState();
                if (source != null && outgoing.ContainsKey(source)) outgoing[source].Add(transition);
                if (target != null && incoming.ContainsKey(target)) incoming[target].Add(transition);

                TransitionPorts ports = GetOrCreateTransitionPorts(transition);
                ports.IsFeedback = source != null && target != null
                    && ranks.ContainsKey(source) && ranks.ContainsKey(target)
                    && ranks[target] <= ranks[source];
            }

            foreach (List<ITransition> stateTransitions in outgoing.Values)
                AssignSourcePorts(stateTransitions, verticalOrder);
            foreach (List<ITransition> stateTransitions in incoming.Values)
                AssignTargetPorts(stateTransitions, verticalOrder);
        }

        private static void AssignSourcePorts(IEnumerable<ITransition> transitions,
            IDictionary<IState, int> verticalOrder)
        {
            List<ITransition> transitionList = transitions.ToList();
            List<ITransition> regularTransitions = transitionList
                .Where(candidate => !GetOrCreateTransitionPorts(candidate).IsFeedback)
                .OrderByDescending(candidate => GetVerticalOrder(candidate.getTargetState(), verticalOrder))
                .ThenBy(candidate => candidate.getModelComponentID())
                .ToList();
            AssignRegularSourcePorts(regularTransitions);

            List<ITransition> feedbackTransitions = transitionList
                .Where(candidate => GetOrCreateTransitionPorts(candidate).IsFeedback)
                .OrderByDescending(candidate => GetVerticalOrder(candidate.getTargetState(), verticalOrder))
                .ThenBy(candidate => candidate.getModelComponentID())
                .ToList();
            for (int index = 0; index < feedbackTransitions.Count; index++)
                GetOrCreateTransitionPorts(feedbackTransitions[index]).SourceY =
                    GetFeedbackPortPosition(index, feedbackTransitions.Count);
        }

        private static void AssignTargetPorts(IEnumerable<ITransition> transitions,
            IDictionary<IState, int> verticalOrder)
        {
            List<ITransition> transitionList = transitions.ToList();
            List<ITransition> regularTransitions = transitionList
                .Where(candidate => !GetOrCreateTransitionPorts(candidate).IsFeedback)
                .OrderByDescending(candidate => GetVerticalOrder(candidate.getSourceState(), verticalOrder))
                .ThenBy(candidate => candidate.getModelComponentID())
                .ToList();
            AssignRegularTargetPorts(regularTransitions);

            List<ITransition> feedbackTransitions = transitionList
                .Where(candidate => GetOrCreateTransitionPorts(candidate).IsFeedback)
                .OrderByDescending(candidate => GetVerticalOrder(candidate.getSourceState(), verticalOrder))
                .ThenBy(candidate => candidate.getModelComponentID())
                .ToList();
            for (int index = 0; index < feedbackTransitions.Count; index++)
                GetOrCreateTransitionPorts(feedbackTransitions[index]).TargetY =
                    GetFeedbackPortPosition(index, feedbackTransitions.Count);
        }

        private static void AssignRegularSourcePorts(IList<ITransition> transitions)
        {
            for (int index = 0; index < transitions.Count; index++)
                GetOrCreateTransitionPorts(transitions[index]).SourceY =
                    GetRegularPortPosition(index, transitions.Count);
        }

        private static void AssignRegularTargetPorts(IList<ITransition> transitions)
        {
            for (int index = 0; index < transitions.Count; index++)
                GetOrCreateTransitionPorts(transitions[index]).TargetY =
                    GetRegularPortPosition(index, transitions.Count);
        }

        private static TransitionPorts GetOrCreateTransitionPorts(ITransition transition)
        {
            string transitionId = transition.getModelComponentID();
            if (!PortsByTransitionId.TryGetValue(transitionId, out TransitionPorts ports))
            {
                ports = new TransitionPorts();
                PortsByTransitionId.Add(transitionId, ports);
            }

            return ports;
        }

        private static void MarkFallbackTransitions(IEnumerable<ITransition> transitions)
        {
            foreach (ITransition transition in transitions)
            {
                TransitionPorts ports = GetOrCreateTransitionPorts(transition);
                ports.UseFallbackRouting = HasGeneratedBounds(transition.getSourceState() as IPASSProcessModelElement)
                    || HasGeneratedBounds(transition.getTargetState() as IPASSProcessModelElement);
            }
        }

        private static double GetRegularPortPosition(int index, int count)
        {
            if (count <= 1) return 0.5;
            return RegularPortMargin
                + index * ((1.0 - 2.0 * RegularPortMargin) / (count - 1.0));
        }

        private static double GetFeedbackPortPosition(int index, int count)
        {
            if (count <= 1) return FeedbackPortPosition;
            return 0.06 + index * (0.10 / (count - 1.0));
        }

        private static Dictionary<IState, int> DetermineStateVerticalOrder(IEnumerable<IState> states,
            IEnumerable<ITransition> transitions, IDictionary<IState, int> ranks)
        {
            List<IState> stateList = states.ToList();
            List<ITransition> transitionList = transitions.ToList();
            Dictionary<IState, int> result = new Dictionary<IState, int>();
            Dictionary<int, List<IState>> statesByRank = stateList
                .GroupBy(state => ranks[state])
                .ToDictionary(group => group.Key,
                    group => group.OrderBy(state => state.getModelComponentID()).ToList());

            foreach (List<IState> rankStates in statesByRank.Values)
                UpdateVerticalOrder(rankStates, result);

            int maxRank = ranks.Count == 0 ? 0 : ranks.Values.Max();
            for (int iteration = 0; iteration < 4; iteration++)
            {
                for (int rank = 1; rank <= maxRank; rank++)
                {
                    if (!statesByRank.TryGetValue(rank, out List<IState> rankStates)) continue;
                    rankStates.Sort((left, right) => CompareByBarycenter(left, right, true,
                        transitionList, ranks, statesByRank, result));
                    UpdateVerticalOrder(rankStates, result);
                }

                for (int rank = maxRank - 1; rank >= 0; rank--)
                {
                    if (!statesByRank.TryGetValue(rank, out List<IState> rankStates)) continue;
                    rankStates.Sort((left, right) => CompareByBarycenter(left, right, false,
                        transitionList, ranks, statesByRank, result));
                    UpdateVerticalOrder(rankStates, result);
                }
            }

            return result;
        }

        private static int CompareByBarycenter(IState left, IState right, bool usePredecessors,
            IEnumerable<ITransition> transitions, IDictionary<IState, int> ranks,
            IDictionary<int, List<IState>> statesByRank, IDictionary<IState, int> verticalOrder)
        {
            double leftBarycenter = GetNeighborBarycenter(left, usePredecessors, transitions, ranks,
                statesByRank, verticalOrder);
            double rightBarycenter = GetNeighborBarycenter(right, usePredecessors, transitions, ranks,
                statesByRank, verticalOrder);
            int comparison = leftBarycenter.CompareTo(rightBarycenter);
            if (comparison != 0) return comparison;

            comparison = verticalOrder[left].CompareTo(verticalOrder[right]);
            return comparison != 0
                ? comparison
                : string.CompareOrdinal(left.getModelComponentID(), right.getModelComponentID());
        }

        private static double GetNeighborBarycenter(IState state, bool usePredecessors,
            IEnumerable<ITransition> transitions, IDictionary<IState, int> ranks,
            IDictionary<int, List<IState>> statesByRank, IDictionary<IState, int> verticalOrder)
        {
            IEnumerable<IState> neighbors = usePredecessors
                ? transitions.Where(transition => transition.getTargetState() == state
                    && transition.getSourceState() != null
                    && ranks.ContainsKey(transition.getSourceState())
                    && ranks[transition.getSourceState()] < ranks[state])
                    .Select(transition => transition.getSourceState())
                : transitions.Where(transition => transition.getSourceState() == state
                    && transition.getTargetState() != null
                    && ranks.ContainsKey(transition.getTargetState())
                    && ranks[transition.getTargetState()] > ranks[state])
                    .Select(transition => transition.getTargetState());

            List<IState> neighborList = neighbors.Distinct().ToList();
            if (neighborList.Count == 0)
                return GetNormalizedVerticalOrder(state, ranks, statesByRank, verticalOrder);

            return neighborList.Average(neighbor =>
                GetNormalizedVerticalOrder(neighbor, ranks, statesByRank, verticalOrder));
        }

        private static double GetNormalizedVerticalOrder(IState state, IDictionary<IState, int> ranks,
            IDictionary<int, List<IState>> statesByRank, IDictionary<IState, int> verticalOrder)
        {
            int count = statesByRank[ranks[state]].Count;
            return (verticalOrder[state] + 0.5) / Math.Max(1, count);
        }

        private static void UpdateVerticalOrder(IEnumerable<IState> states,
            IDictionary<IState, int> verticalOrder)
        {
            int index = 0;
            foreach (IState state in states)
                verticalOrder[state] = index++;
        }

        private static int GetVerticalOrder(IState state, IDictionary<IState, int> verticalOrder)
        {
            return state != null && verticalOrder.TryGetValue(state, out int order) ? order : 0;
        }

        private static double GetRequiredStateHeight(IState state, IEnumerable<ITransition> transitions)
        {
            int outgoingRegularCount = transitions.Count(transition => transition.getSourceState() == state
                && !GetOrCreateTransitionPorts(transition).IsFeedback);
            int incomingRegularCount = transitions.Count(transition => transition.getTargetState() == state
                && !GetOrCreateTransitionPorts(transition).IsFeedback);
            bool hasOutgoingFeedback = transitions.Any(transition => transition.getSourceState() == state
                && GetOrCreateTransitionPorts(transition).IsFeedback);
            bool hasIncomingFeedback = transitions.Any(transition => transition.getTargetState() == state
                && GetOrCreateTransitionPorts(transition).IsFeedback);
            int outgoingCount = outgoingRegularCount + (hasOutgoingFeedback ? 1 : 0);
            int incomingCount = incomingRegularCount + (hasIncomingFeedback ? 1 : 0);
            int regularPortCount = Math.Max(outgoingCount, incomingCount);
            return Math.Min(0.11, StateHeight + Math.Max(0, regularPortCount - 1) * StateConnectionPitch);
        }

        private static void ScaleRowsToAvailableHeight(IList<double> heights, ref double rowGap)
        {
            if (heights.Count == 0) return;

            double availableHeight = 1.0 - 2.0 * Margin;
            double gaps = Math.Max(0, heights.Count - 1) * StateRowGap;
            double totalHeight = heights.Sum();
            double totalRequired = totalHeight + gaps;
            if (totalRequired <= availableHeight) return;

            double scale = availableHeight / totalRequired;
            rowGap *= scale;
            for (int index = 0; index < heights.Count; index++)
                heights[index] *= scale;
        }

        private static Dictionary<IState, int> DetermineStateRanks(IEnumerable<IState> states, IEnumerable<ITransition> transitions)
        {
            List<IState> stateList = states.ToList();
            Dictionary<IState, int> ranks = stateList.ToDictionary(state => state, state => 0);
            Dictionary<IState, int> incomingCounts = stateList.ToDictionary(state => state, state => 0);
            Dictionary<IState, List<IState>> successors = stateList.ToDictionary(state => state, state => new List<IState>());

            foreach (ITransition transition in transitions)
            {
                IState source = transition.getSourceState();
                IState target = transition.getTargetState();
                if (source != null && target != null && successors.ContainsKey(source) && incomingCounts.ContainsKey(target))
                {
                    successors[source].Add(target);
                    incomingCounts[target]++;
                }
            }

            Queue<IState> queue = new Queue<IState>(incomingCounts
                .Where(pair => pair.Value == 0)
                .Select(pair => pair.Key)
                .OrderBy(candidate => candidate.getModelComponentID()));
            HashSet<IState> processed = new HashSet<IState>();
            while (processed.Count < stateList.Count)
            {
                if (queue.Count == 0)
                {
                    IState cycleRoot = stateList
                        .Where(candidate => !processed.Contains(candidate))
                        .OrderBy(candidate => candidate.getModelComponentID())
                        .First();
                    queue.Enqueue(cycleRoot);
                }

                IState currentState = queue.Dequeue();
                if (!processed.Add(currentState)) continue;
                foreach (IState target in successors[currentState])
                {
                    if (processed.Contains(target)) continue;

                    ranks[target] = Math.Max(ranks[target], ranks[currentState] + 1);
                    incomingCounts[target]--;
                    if (incomingCounts[target] == 0) queue.Enqueue(target);
                }
            }

            return ranks;
        }

        private static int GetMessageExchangeCount(ISubject subject)
        {
            return subject.getIncomingMessageExchanges().Count + subject.getOutgoingMessageExchanges().Count;
        }

        private static bool NeedsFallbackBounds(IVisioExportableWithShape exportable)
        {
            IPASSProcessModelElement element = exportable as IPASSProcessModelElement;
            return exportable != null && element != null && !HasGeneratedBounds(element)
                && !PrepareUsableBounds(exportable, element);
        }

        internal static bool TryGetGeneratedBounds(IPASSProcessModelElement element, out List<ISimple2DVisualizationPoint> bounds)
        {
            if (element != null && GeneratedBounds.TryGetValue(element, out LayoutBounds generatedBounds))
            {
                bounds = new List<ISimple2DVisualizationPoint>
                {
                    CreatePoint(generatedBounds.X, generatedBounds.Y),
                    CreatePoint(generatedBounds.Width, generatedBounds.Height)
                };
                return true;
            }

            bounds = null;
            return false;
        }

        private static bool HasGeneratedBounds(IPASSProcessModelElement element)
        {
            return element != null && GeneratedBounds.ContainsKey(element);
        }

        private static bool PrepareUsableBounds(IVisioExportableWithShape exportable, IPASSProcessModelElement element)
        {
            if (HasUsableBounds(element)) return true;
            if (!exportable.PrepareDimensions()) return false;
            return HasUsableBounds(element);
        }

        private static bool HasUsableBounds(IPASSProcessModelElement element)
        {
            List<ISimple2DVisualizationPoint> bounds = element.getElementsWithUnspecifiedRelation().Values
                .OfType<ISimple2DVisualizationPoint>()
                .ToList();
            if (bounds.Count < 2) return false;

            double x = bounds[0].getRelative2DPosX();
            double y = bounds[0].getRelative2DPosY();
            double width = bounds[1].getRelative2DPosX();
            double height = bounds[1].getRelative2DPosY();

            // The API represents missing data with a full-page default box.
            // A real element must have a positive, smaller-than-page size.
            return IsFinite(x) && IsFinite(y) && IsFinite(width) && IsFinite(height)
                && x >= 0 && x <= 1 && y >= 0 && y <= 1
                && width > 0 && width < 1 && height > 0 && height < 1;
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static double GetRankPosition(int rank, int maxRank, double horizontalMargin)
        {
            return maxRank == 0 ? 0.5 : horizontalMargin + rank * ((1.0 - 2 * horizontalMargin) / maxRank);
        }

        private static double GetRowPosition(int row, int count)
        {
            return 1.0 - Margin - (row + 1) * ((1.0 - 2 * Margin) / (count + 1));
        }

        private static double GetNodeWidth(double maximumWidth, int maxRank, double horizontalMargin)
        {
            return Math.Min(maximumWidth, (1.0 - 2 * horizontalMargin) / (maxRank + 1) * 0.65);
        }

        private static double GetNodeHeight(double maximumHeight, int rows)
        {
            return Math.Min(maximumHeight, (1.0 - 2 * Margin) / (rows + 1) * 0.65);
        }

        private static void SetBounds(IPASSProcessModelElement element, double x, double y, double width, double height)
        {
            GeneratedBounds[element] = new LayoutBounds(x, y, width, height);
        }

        private static Simple2DVisualizationPoint CreatePoint(double x, double y)
        {
            Simple2DVisualizationPoint point = new Simple2DVisualizationPoint();
            point.setRelative2DPosX(x);
            point.setRelative2DPosY(y);
            return point;
        }

        private sealed class LayoutBounds
        {
            public LayoutBounds(double x, double y, double width, double height)
            {
                X = x;
                Y = y;
                Width = width;
                Height = height;
            }

            public double X { get; private set; }
            public double Y { get; private set; }
            public double Width { get; private set; }
            public double Height { get; private set; }
        }

        private sealed class TransitionPorts
        {
            public double SourceY { get; set; } = 0.5;
            public double TargetY { get; set; } = 0.5;
            public bool IsFeedback { get; set; }
            public bool UseFallbackRouting { get; set; }
        }
    }
}
