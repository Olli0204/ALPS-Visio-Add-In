using alps.net.api.StandardPASS;
using ALPS_Visio_AddIn_rewrite.OWLShapes.Layout;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System.Collections.Generic;
using System.Linq;

namespace ALPS_Visio_AddIn_rewrite.Tests.Layout
{
    [TestClass]
    public sealed class FallbackLayoutTests
    {
        [TestMethod]
        public void DetermineStateRanks_AcyclicGraph_AssignsLongestPathRanks()
        {
            Mock<IState> start = State("A");
            Mock<IState> left = State("B");
            Mock<IState> right = State("C");
            Mock<IState> end = State("D");
            ITransition[] transitions =
            {
                Transition("A-B", start.Object, left.Object).Object,
                Transition("A-C", start.Object, right.Object).Object,
                Transition("C-D", right.Object, end.Object).Object
            };

            Dictionary<IState, int> ranks =
                FallbackLayoutRanker.DetermineStateRanks(
                    new[]
                    {
                        end.Object, right.Object,
                        left.Object, start.Object
                    },
                    transitions);

            Assert.AreEqual(0, ranks[start.Object]);
            Assert.AreEqual(1, ranks[left.Object]);
            Assert.AreEqual(1, ranks[right.Object]);
            Assert.AreEqual(2, ranks[end.Object]);
        }

        [TestMethod]
        public void DetermineStateRanks_Cycle_UsesStableComponentIdRoot()
        {
            Mock<IState> beta = State("B");
            Mock<IState> alpha = State("A");
            ITransition[] transitions =
            {
                Transition("A-B", alpha.Object, beta.Object).Object,
                Transition("B-A", beta.Object, alpha.Object).Object
            };

            Dictionary<IState, int> ranks =
                FallbackLayoutRanker.DetermineStateRanks(
                    new[] { beta.Object, alpha.Object }, transitions);

            Assert.AreEqual(0, ranks[alpha.Object]);
            Assert.AreEqual(1, ranks[beta.Object]);
        }

        [TestMethod]
        public void DetermineStateVerticalOrder_InputOrderChanges_ReturnsSameOrder()
        {
            Mock<IState> start = State("A");
            Mock<IState> beta = State("B");
            Mock<IState> charlie = State("C");
            ITransition[] transitions =
            {
                Transition("A-B", start.Object, beta.Object).Object,
                Transition("A-C", start.Object, charlie.Object).Object
            };
            Dictionary<IState, int> ranks = new Dictionary<IState, int>
            {
                [start.Object] = 0,
                [beta.Object] = 1,
                [charlie.Object] = 1
            };

            Dictionary<IState, int> forward =
                FallbackLayoutRanker.DetermineStateVerticalOrder(
                    new[] { start.Object, beta.Object, charlie.Object },
                    transitions,
                    ranks);
            Dictionary<IState, int> reversed =
                FallbackLayoutRanker.DetermineStateVerticalOrder(
                    new[] { charlie.Object, beta.Object, start.Object },
                    transitions.Reverse(),
                    ranks);

            Assert.AreEqual(forward[beta.Object], reversed[beta.Object]);
            Assert.AreEqual(forward[charlie.Object], reversed[charlie.Object]);
            Assert.AreNotEqual(forward[beta.Object], forward[charlie.Object]);
        }

        [TestMethod]
        public void Prepare_MultipleOutgoingTransitions_SpreadsPortsDeterministically()
        {
            Mock<IState> start = State("A");
            Mock<IState> upper = State("B");
            Mock<IState> lower = State("C");
            Mock<ITransition> toUpper =
                Transition("A-B", start.Object, upper.Object);
            Mock<ITransition> toLower =
                Transition("A-C", start.Object, lower.Object);
            Mock<ITransition> feedback =
                Transition("B-A", upper.Object, start.Object);
            FallbackLayoutState layoutState = new FallbackLayoutState();

            FallbackTransitionPortPlanner.Prepare(
                new[] { start.Object, upper.Object, lower.Object },
                new[] { toUpper.Object, toLower.Object, feedback.Object },
                new Dictionary<IState, int>
                {
                    [start.Object] = 0,
                    [upper.Object] = 1,
                    [lower.Object] = 1
                },
                new Dictionary<IState, int>
                {
                    [start.Object] = 0,
                    [upper.Object] = 0,
                    [lower.Object] = 1
                },
                layoutState);

            Assert.IsTrue(layoutState.TryGetTransitionPorts(
                toUpper.Object, out FallbackLayoutState.TransitionPorts upperPorts));
            Assert.IsTrue(layoutState.TryGetTransitionPorts(
                toLower.Object, out FallbackLayoutState.TransitionPorts lowerPorts));
            Assert.IsTrue(layoutState.TryGetTransitionPorts(
                feedback.Object, out FallbackLayoutState.TransitionPorts feedbackPorts));
            Assert.AreEqual(0.75, upperPorts.SourceY, 0.0001);
            Assert.AreEqual(0.25, lowerPorts.SourceY, 0.0001);
            Assert.IsTrue(feedbackPorts.IsFeedback);
        }

        [TestMethod]
        public void ResetBounds_PreviousLayoutPass_RemovesStoredBounds()
        {
            Mock<IState> state = State("A");
            FallbackLayoutState layoutState = new FallbackLayoutState();
            layoutState.SetBounds(state.Object, 1, 2, 3, 4);

            layoutState.ResetBounds();

            Assert.IsFalse(layoutState.HasBounds(state.Object));
        }

        [TestMethod]
        public void MarkFallbackTransitions_GeneratedEndpoint_MarksRouting()
        {
            Mock<IState> source = State("A");
            Mock<IState> target = State("B");
            Mock<ITransition> transition =
                Transition("A-B", source.Object, target.Object);
            FallbackLayoutState layoutState = new FallbackLayoutState();

            FallbackTransitionPortPlanner.MarkFallbackTransitions(
                new[] { transition.Object },
                layoutState,
                element => ReferenceEquals(element, target.Object));

            Assert.IsTrue(layoutState.TryGetTransitionPorts(
                transition.Object, out FallbackLayoutState.TransitionPorts ports));
            Assert.IsTrue(ports.UseFallbackRouting);
        }

        private static Mock<IState> State(string id)
        {
            Mock<IState> state = new Mock<IState>();
            state.Setup(candidate => candidate.getModelComponentID())
                .Returns(id);
            return state;
        }

        private static Mock<ITransition> Transition(
            string id, IState source, IState target)
        {
            Mock<ITransition> transition = new Mock<ITransition>();
            transition.Setup(candidate => candidate.getModelComponentID())
                .Returns(id);
            transition.Setup(candidate => candidate.getSourceState())
                .Returns(source);
            transition.Setup(candidate => candidate.getTargetState())
                .Returns(target);
            return transition;
        }
    }
}
