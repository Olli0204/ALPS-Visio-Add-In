using alps.net.api.ALPS;
using alps.net.api.StandardPASS;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes.Layout
{
    /// <summary>
    /// Calculates deterministic graph ranks and row order without Visio or COM state.
    /// </summary>
    internal static class FallbackLayoutRanker
    {
        public static Dictionary<ISubject, int> DetermineSubjectRanks(
            IEnumerable<ISubject> subjects)
        {
            List<ISubject> subjectList = subjects.ToList();
            Dictionary<ISubject, int> ranks =
                subjectList.ToDictionary(subject => subject, subject => 0);
            Dictionary<ISubject, int> incomingCounts =
                subjectList.ToDictionary(subject => subject, subject => 0);

            foreach (ISubject subject in subjectList)
            {
                foreach (IMessageExchange exchange
                    in subject.getOutgoingMessageExchanges().Values)
                {
                    if (exchange.getReceiver() != null
                        && incomingCounts.ContainsKey(exchange.getReceiver()))
                    {
                        incomingCounts[exchange.getReceiver()]++;
                    }
                }
            }

            Queue<ISubject> queue = new Queue<ISubject>(incomingCounts
                .Where(pair => pair.Value == 0)
                .Select(pair => pair.Key)
                .OrderBy(candidate => candidate.getModelComponentID()));
            HashSet<ISubject> processed = new HashSet<ISubject>();
            while (processed.Count < subjectList.Count)
            {
                // A bidirectional exchange has no natural root. Pick a stable
                // one so connected subjects receive separate columns.
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
                if (!processed.Add(subject))
                    continue;

                foreach (IMessageExchange exchange
                    in subject.getOutgoingMessageExchanges().Values)
                {
                    ISubject receiver = exchange.getReceiver();
                    if (receiver == null || processed.Contains(receiver)
                        || !incomingCounts.ContainsKey(receiver))
                    {
                        continue;
                    }

                    ranks[receiver] = Math.Max(ranks[receiver], ranks[subject] + 1);
                    incomingCounts[receiver]--;
                    if (incomingCounts[receiver] == 0)
                        queue.Enqueue(receiver);
                }
            }

            return ranks;
        }

        public static Dictionary<IState, int> DetermineStateRanks(
            IEnumerable<IState> states, IEnumerable<ITransition> transitions)
        {
            List<IState> stateList = states.ToList();
            Dictionary<IState, int> ranks =
                stateList.ToDictionary(state => state, state => 0);
            Dictionary<IState, int> incomingCounts =
                stateList.ToDictionary(state => state, state => 0);
            Dictionary<IState, List<IState>> successors =
                stateList.ToDictionary(state => state, state => new List<IState>());

            foreach (ITransition transition in transitions)
            {
                IState source = transition.getSourceState();
                IState target = transition.getTargetState();
                if (source != null && target != null
                    && successors.ContainsKey(source)
                    && incomingCounts.ContainsKey(target))
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
                if (!processed.Add(currentState))
                    continue;

                foreach (IState target in successors[currentState])
                {
                    if (processed.Contains(target))
                        continue;

                    ranks[target] = Math.Max(
                        ranks[target], ranks[currentState] + 1);
                    incomingCounts[target]--;
                    if (incomingCounts[target] == 0)
                        queue.Enqueue(target);
                }
            }

            return ranks;
        }

        public static Dictionary<IState, int> DetermineStateVerticalOrder(
            IEnumerable<IState> states, IEnumerable<ITransition> transitions,
            IDictionary<IState, int> ranks)
        {
            List<IState> stateList = states.ToList();
            List<ITransition> transitionList = transitions.ToList();
            Dictionary<IState, int> result = new Dictionary<IState, int>();
            Dictionary<int, List<IState>> statesByRank = stateList
                .GroupBy(state => ranks[state])
                .ToDictionary(
                    group => group.Key,
                    group => group
                        .OrderBy(state => state.getModelComponentID())
                        .ToList());

            foreach (List<IState> rankStates in statesByRank.Values)
                UpdateVerticalOrder(rankStates, result);

            int maxRank = ranks.Count == 0 ? 0 : ranks.Values.Max();
            for (int iteration = 0; iteration < 4; iteration++)
            {
                for (int rank = 1; rank <= maxRank; rank++)
                {
                    if (!statesByRank.TryGetValue(rank, out List<IState> rankStates))
                        continue;

                    rankStates.Sort((left, right) => CompareByBarycenter(
                        left, right, true, transitionList, ranks, statesByRank, result));
                    UpdateVerticalOrder(rankStates, result);
                }

                for (int rank = maxRank - 1; rank >= 0; rank--)
                {
                    if (!statesByRank.TryGetValue(rank, out List<IState> rankStates))
                        continue;

                    rankStates.Sort((left, right) => CompareByBarycenter(
                        left, right, false, transitionList, ranks, statesByRank, result));
                    UpdateVerticalOrder(rankStates, result);
                }
            }

            return result;
        }

        public static int GetMessageExchangeCount(ISubject subject)
        {
            return subject.getIncomingMessageExchanges().Count
                + subject.getOutgoingMessageExchanges().Count;
        }

        private static int CompareByBarycenter(IState left, IState right,
            bool usePredecessors, IEnumerable<ITransition> transitions,
            IDictionary<IState, int> ranks,
            IDictionary<int, List<IState>> statesByRank,
            IDictionary<IState, int> verticalOrder)
        {
            double leftBarycenter = GetNeighborBarycenter(
                left, usePredecessors, transitions, ranks, statesByRank, verticalOrder);
            double rightBarycenter = GetNeighborBarycenter(
                right, usePredecessors, transitions, ranks, statesByRank, verticalOrder);
            int comparison = leftBarycenter.CompareTo(rightBarycenter);
            if (comparison != 0)
                return comparison;

            comparison = verticalOrder[left].CompareTo(verticalOrder[right]);
            return comparison != 0
                ? comparison
                : string.CompareOrdinal(
                    left.getModelComponentID(), right.getModelComponentID());
        }

        private static double GetNeighborBarycenter(IState state,
            bool usePredecessors, IEnumerable<ITransition> transitions,
            IDictionary<IState, int> ranks,
            IDictionary<int, List<IState>> statesByRank,
            IDictionary<IState, int> verticalOrder)
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
            {
                return GetNormalizedVerticalOrder(
                    state, ranks, statesByRank, verticalOrder);
            }

            return neighborList.Average(neighbor => GetNormalizedVerticalOrder(
                neighbor, ranks, statesByRank, verticalOrder));
        }

        private static double GetNormalizedVerticalOrder(IState state,
            IDictionary<IState, int> ranks,
            IDictionary<int, List<IState>> statesByRank,
            IDictionary<IState, int> verticalOrder)
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
    }
}
