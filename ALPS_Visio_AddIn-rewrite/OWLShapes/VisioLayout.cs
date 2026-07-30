using alps.net.api.ALPS;
using alps.net.api.StandardPASS;
using System;
using System.Collections.Generic;
using System.Linq;
using ALPS_Visio_AddIn_rewrite.OWLShapes.Layout;
using LayoutBounds =
    ALPS_Visio_AddIn_rewrite.OWLShapes.Layout.FallbackLayoutState.LayoutBounds;
using TransitionPorts =
    ALPS_Visio_AddIn_rewrite.OWLShapes.Layout.FallbackLayoutState.TransitionPorts;

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
        private const double PortMargin = 0.25;
        private static readonly FallbackLayoutState LayoutState =
            new FallbackLayoutState();

        public static bool ArrangeModelLayer(IEnumerable<IPASSProcessModelElement> elements)
        {
            LayoutState.ResetBounds();
            List<ISubject> subjects = elements.OfType<ISubject>().ToList();
            Dictionary<ISubject, int> ranks =
                FallbackLayoutRanker.DetermineSubjectRanks(subjects);
            return ArrangeSubjects(subjects, ranks);
        }

        public static bool ArrangeBehavior(IEnumerable<IBehaviorDescribingComponent> components)
        {
            List<IState> states = components.OfType<IState>().ToList();
            List<ITransition> transitions = components.OfType<ITransition>().ToList();
            LayoutState.ResetTransitionPorts();
            Dictionary<IState, int> ranks =
                FallbackLayoutRanker.DetermineStateRanks(states, transitions);
            Dictionary<IState, int> verticalOrder =
                FallbackLayoutRanker.DetermineStateVerticalOrder(
                    states, transitions, ranks);
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

            if (LayoutState.TryGetTransitionPorts(
                    transition, out TransitionPorts ports)
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
                    .OrderByDescending(FallbackLayoutRanker.GetMessageExchangeCount)
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
            List<ITransition> sortedTransitions = transitions
                .OrderBy(candidate => GetOrCreateTransitionPorts(candidate).IsFeedback ? 0 : 1)
                .ThenByDescending(candidate => GetVerticalOrder(candidate.getTargetState(), verticalOrder))
                .ThenBy(candidate => candidate.getModelComponentID())
                .ToList();
            for (int index = 0; index < sortedTransitions.Count; index++)
                GetOrCreateTransitionPorts(sortedTransitions[index]).SourceY =
                    GetPortPosition(index, sortedTransitions.Count);
        }

        private static void AssignTargetPorts(IEnumerable<ITransition> transitions,
            IDictionary<IState, int> verticalOrder)
        {
            List<ITransition> sortedTransitions = transitions
                .OrderBy(candidate => GetOrCreateTransitionPorts(candidate).IsFeedback ? 0 : 1)
                .ThenByDescending(candidate => GetVerticalOrder(candidate.getSourceState(), verticalOrder))
                .ThenBy(candidate => candidate.getModelComponentID())
                .ToList();
            for (int index = 0; index < sortedTransitions.Count; index++)
                GetOrCreateTransitionPorts(sortedTransitions[index]).TargetY =
                    GetPortPosition(index, sortedTransitions.Count);
        }

        private static TransitionPorts GetOrCreateTransitionPorts(ITransition transition)
        {
            return LayoutState.GetOrCreateTransitionPorts(transition);
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

        private static double GetPortPosition(int index, int count)
        {
            if (count <= 1) return 0.5;
            return PortMargin
                + index * ((1.0 - 2.0 * PortMargin) / (count - 1.0));
        }

        private static int GetVerticalOrder(IState state, IDictionary<IState, int> verticalOrder)
        {
            return state != null && verticalOrder.TryGetValue(state, out int order) ? order : 0;
        }

        private static double GetRequiredStateHeight(IState state, IEnumerable<ITransition> transitions)
        {
            int outgoingCount = transitions.Count(transition => transition.getSourceState() == state);
            int incomingCount = transitions.Count(transition => transition.getTargetState() == state);
            int portCount = Math.Max(outgoingCount, incomingCount);
            return Math.Min(0.11, StateHeight + Math.Max(0, portCount - 1) * StateConnectionPitch);
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

        private static bool NeedsFallbackBounds(IVisioExportableWithShape exportable)
        {
            IPASSProcessModelElement element = exportable as IPASSProcessModelElement;
            return exportable != null && element != null && !HasGeneratedBounds(element)
                && !PrepareUsableBounds(exportable, element);
        }

        internal static bool TryGetGeneratedBounds(IPASSProcessModelElement element, out List<ISimple2DVisualizationPoint> bounds)
        {
            if (LayoutState.TryGetBounds(element, out LayoutBounds generatedBounds))
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
            return LayoutState.HasBounds(element);
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
            LayoutState.SetBounds(element, x, y, width, height);
        }

        private static Simple2DVisualizationPoint CreatePoint(double x, double y)
        {
            Simple2DVisualizationPoint point = new Simple2DVisualizationPoint();
            point.setRelative2DPosX(x);
            point.setRelative2DPosY(y);
            return point;
        }

    }
}
