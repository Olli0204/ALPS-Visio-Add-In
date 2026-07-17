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
        private const double SubjectWidth = 0.22;
        private const double SubjectHeight = 0.16;
        private const double StateWidth = 0.18;
        private const double StateHeight = 0.12;

        public static void ArrangeModelLayer(IEnumerable<IPASSProcessModelElement> elements)
        {
            List<ISubject> subjects = elements.OfType<ISubject>().ToList();
            Dictionary<ISubject, int> ranks = DetermineSubjectRanks(subjects);
            ArrangeSubjects(subjects, ranks);
        }

        public static void ArrangeBehavior(IEnumerable<IBehaviorDescribingComponent> components)
        {
            List<IState> states = components.OfType<IState>().ToList();
            Dictionary<IState, int> ranks = DetermineStateRanks(states, components.OfType<ITransition>());
            ArrangeStates(states, ranks);
        }

        public static void PrepareOrArrange(IVisioExportableWithShape exportable, int index, bool subjectDiagram)
        {
            IPASSProcessModelElement element = exportable as IPASSProcessModelElement;
            if (exportable == null || element == null || HasBounds(element)) return;
            if (exportable.PrepareDimensions() && HasBounds(element)) return;

            int columns = subjectDiagram ? 4 : 3;
            int column = index % columns;
            int row = index / columns;
            double x = Margin + column * ((1.0 - 2 * Margin) / Math.Max(1, columns - 1));
            double y = Math.Max(Margin, 1.0 - Margin - row * 0.20);
            SetBounds(element, x, y, subjectDiagram ? StateWidth : SubjectWidth,
                subjectDiagram ? StateHeight : SubjectHeight);
        }

        private static void ArrangeSubjects(IEnumerable<ISubject> subjects, IDictionary<ISubject, int> ranks)
        {
            int maxRank = ranks.Count == 0 ? 0 : ranks.Values.Max();
            double width = GetNodeWidth(SubjectWidth, maxRank);
            foreach (IGrouping<int, ISubject> rankGroup in subjects.GroupBy(subject => ranks[subject]).OrderBy(group => group.Key))
            {
                List<ISubject> missingSubjects = rankGroup
                    .Where(subject => NeedsFallbackBounds(subject as IVisioExportableWithShape))
                    .OrderByDescending(GetMessageExchangeCount)
                    .ThenBy(subject => subject.getModelComponentID())
                    .ToList();

                for (int row = 0; row < missingSubjects.Count; row++)
                {
                    double x = GetRankPosition(rankGroup.Key, maxRank);
                    double y = GetRowPosition(row, missingSubjects.Count);
                    SetBounds((IPASSProcessModelElement)missingSubjects[row], x, y, width,
                        GetNodeHeight(SubjectHeight, missingSubjects.Count));
                }
            }
        }

        private static void ArrangeStates(IEnumerable<IState> states, IDictionary<IState, int> ranks)
        {
            int maxRank = ranks.Count == 0 ? 0 : ranks.Values.Max();
            double width = GetNodeWidth(StateWidth, maxRank);
            foreach (IGrouping<int, IState> rankGroup in states.GroupBy(state => ranks[state]).OrderBy(group => group.Key))
            {
                List<IState> missingStates = rankGroup
                    .Where(state => NeedsFallbackBounds(state as IVisioExportableWithShape))
                    .OrderBy(state => state.getModelComponentID())
                    .ToList();

                for (int row = 0; row < missingStates.Count; row++)
                {
                    double x = GetRankPosition(rankGroup.Key, maxRank);
                    double y = GetRowPosition(row, missingStates.Count);
                    SetBounds((IPASSProcessModelElement)missingStates[row], x, y, width,
                        GetNodeHeight(StateHeight, missingStates.Count));
                }
            }
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

            Queue<ISubject> queue = new Queue<ISubject>(incomingCounts.Where(pair => pair.Value == 0).Select(pair => pair.Key));
            while (queue.Count > 0)
            {
                ISubject subject = queue.Dequeue();
                foreach (IMessageExchange exchange in subject.getOutgoingMessageExchanges().Values)
                {
                    ISubject receiver = exchange.getReceiver();
                    if (receiver == null || !incomingCounts.ContainsKey(receiver)) continue;

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

            Queue<IState> queue = new Queue<IState>(incomingCounts.Where(pair => pair.Value == 0).Select(pair => pair.Key));
            while (queue.Count > 0)
            {
                IState state = queue.Dequeue();
                foreach (IState target in successors[state])
                {
                    ranks[target] = Math.Max(ranks[target], ranks[state] + 1);
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
            if (exportable == null || element == null || HasBounds(element)) return false;
            return !exportable.PrepareDimensions() || !HasBounds(element);
        }

        private static bool HasBounds(IPASSProcessModelElement element)
        {
            return element.getElementsWithUnspecifiedRelation().Values.OfType<ISimple2DVisualizationPoint>().Count() >= 2;
        }

        private static double GetRankPosition(int rank, int maxRank)
        {
            return maxRank == 0 ? 0.5 : Margin + rank * ((1.0 - 2 * Margin) / maxRank);
        }

        private static double GetRowPosition(int row, int count)
        {
            return 1.0 - Margin - (row + 1) * ((1.0 - 2 * Margin) / (count + 1));
        }

        private static double GetNodeWidth(double maximumWidth, int maxRank)
        {
            return Math.Min(maximumWidth, (1.0 - 2 * Margin) / (maxRank + 1) * 0.65);
        }

        private static double GetNodeHeight(double maximumHeight, int rows)
        {
            return Math.Min(maximumHeight, (1.0 - 2 * Margin) / (rows + 1) * 0.65);
        }

        private static void SetBounds(IPASSProcessModelElement element, double x, double y, double width, double height)
        {
            AddPoint(element, x, y);
            AddPoint(element, width, height);
        }

        private static void AddPoint(IPASSProcessModelElement element, double x, double y)
        {
            Simple2DVisualizationPoint point = new Simple2DVisualizationPoint();
            point.setRelative2DPosX(x);
            point.setRelative2DPosY(y);
            element.addElementWithUnspecifiedRelation(point);
        }
    }
}
