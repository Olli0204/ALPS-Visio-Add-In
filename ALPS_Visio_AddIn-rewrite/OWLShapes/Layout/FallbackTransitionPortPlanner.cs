using alps.net.api.StandardPASS;
using System;
using System.Collections.Generic;
using System.Linq;
using TransitionPorts =
    ALPS_Visio_AddIn_rewrite.OWLShapes.Layout.FallbackLayoutState.TransitionPorts;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes.Layout
{
    /// <summary>
    /// Assigns deterministic connector ports and marks transitions that need
    /// fallback routing.
    /// </summary>
    internal static class FallbackTransitionPortPlanner
    {
        private const double PortMargin = 0.25;

        public static void Prepare(IEnumerable<IState> states,
            IEnumerable<ITransition> transitions, IDictionary<IState, int> ranks,
            IDictionary<IState, int> verticalOrder, FallbackLayoutState state)
        {
            List<IState> stateList = states.ToList();
            Dictionary<IState, List<ITransition>> outgoing =
                stateList.ToDictionary(
                    candidate => candidate, candidate => new List<ITransition>());
            Dictionary<IState, List<ITransition>> incoming =
                stateList.ToDictionary(
                    candidate => candidate, candidate => new List<ITransition>());

            foreach (ITransition transition in transitions)
            {
                IState source = transition.getSourceState();
                IState target = transition.getTargetState();
                if (source != null && outgoing.ContainsKey(source))
                    outgoing[source].Add(transition);
                if (target != null && incoming.ContainsKey(target))
                    incoming[target].Add(transition);

                TransitionPorts ports = state.GetOrCreateTransitionPorts(transition);
                ports.IsFeedback = source != null && target != null
                    && ranks.ContainsKey(source) && ranks.ContainsKey(target)
                    && ranks[target] <= ranks[source];
            }

            foreach (List<ITransition> stateTransitions in outgoing.Values)
                AssignSourcePorts(stateTransitions, verticalOrder, state);
            foreach (List<ITransition> stateTransitions in incoming.Values)
                AssignTargetPorts(stateTransitions, verticalOrder, state);
        }

        public static void MarkFallbackTransitions(
            IEnumerable<ITransition> transitions, FallbackLayoutState state,
            Func<IPASSProcessModelElement, bool> hasGeneratedBounds)
        {
            foreach (ITransition transition in transitions)
            {
                TransitionPorts ports =
                    state.GetOrCreateTransitionPorts(transition);
                ports.UseFallbackRouting = hasGeneratedBounds(
                        transition.getSourceState() as IPASSProcessModelElement)
                    || hasGeneratedBounds(
                        transition.getTargetState() as IPASSProcessModelElement);
            }
        }

        private static void AssignSourcePorts(
            IEnumerable<ITransition> transitions,
            IDictionary<IState, int> verticalOrder, FallbackLayoutState state)
        {
            List<ITransition> sortedTransitions = transitions
                .OrderBy(candidate =>
                    state.GetOrCreateTransitionPorts(candidate).IsFeedback ? 0 : 1)
                .ThenByDescending(candidate =>
                    GetVerticalOrder(candidate.getTargetState(), verticalOrder))
                .ThenBy(candidate => candidate.getModelComponentID())
                .ToList();

            for (int index = 0; index < sortedTransitions.Count; index++)
            {
                state.GetOrCreateTransitionPorts(sortedTransitions[index]).SourceY =
                    GetPortPosition(index, sortedTransitions.Count);
            }
        }

        private static void AssignTargetPorts(
            IEnumerable<ITransition> transitions,
            IDictionary<IState, int> verticalOrder, FallbackLayoutState state)
        {
            List<ITransition> sortedTransitions = transitions
                .OrderBy(candidate =>
                    state.GetOrCreateTransitionPorts(candidate).IsFeedback ? 0 : 1)
                .ThenByDescending(candidate =>
                    GetVerticalOrder(candidate.getSourceState(), verticalOrder))
                .ThenBy(candidate => candidate.getModelComponentID())
                .ToList();

            for (int index = 0; index < sortedTransitions.Count; index++)
            {
                state.GetOrCreateTransitionPorts(sortedTransitions[index]).TargetY =
                    GetPortPosition(index, sortedTransitions.Count);
            }
        }

        private static int GetVerticalOrder(IState state,
            IDictionary<IState, int> verticalOrder)
        {
            return state != null
                && verticalOrder.TryGetValue(state, out int order)
                ? order
                : 0;
        }

        private static double GetPortPosition(int index, int count)
        {
            if (count <= 1)
                return 0.5;

            return PortMargin
                + index * ((1.0 - 2.0 * PortMargin) / (count - 1.0));
        }
    }
}
