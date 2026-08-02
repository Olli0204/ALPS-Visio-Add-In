using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAddIn.Snapping;

namespace ALPS_Visio_AddIn_rewrite.Tests.Snapping
{
    [TestClass]
    public sealed class SnapHandlerTests
    {
        [TestMethod]
        public void IsWithinSnapRange_PointOnCircularBoundary_ReturnsTrue()
        {
            Assert.IsTrue(SnapHandler.IsWithinSnapRange(12d, 16d, 20d));
        }

        [TestMethod]
        public void IsWithinSnapRange_PointOutsideCircularBoundary_ReturnsFalse()
        {
            Assert.IsFalse(SnapHandler.IsWithinSnapRange(15d, 15d, 20d));
        }

        [TestMethod]
        public void IsWithinSnapRange_SameXAndBoundaryY_ReturnsTrue()
        {
            Assert.IsTrue(SnapHandler.IsWithinSnapRange(0d, 20d, 20d));
        }

        [TestMethod]
        public void IsWithinSnapRange_NegativeRange_Throws()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                SnapHandler.IsWithinSnapRange(0d, 0d, -1d));
        }

        [TestMethod]
        public void IsActorExtensionIdentity_GuardExtensionMaster_ReturnsTrue()
        {
            Assert.IsTrue(SidSnapHandler.IsActorExtensionIdentity(
                false, "GuardExtension", null));
        }

        [TestMethod]
        public void IsActorExtensionIdentity_GuardExtensionInstance_ReturnsTrue()
        {
            Assert.IsTrue(SidSnapHandler.IsActorExtensionIdentity(
                false, null, "GuardExtension.42"));
        }

        [TestMethod]
        public void IsActorExtensionIdentity_StandardActor_ReturnsFalse()
        {
            Assert.IsFalse(SidSnapHandler.IsActorExtensionIdentity(
                false, "StandardActor", "StandardActor.42"));
        }

        [TestMethod]
        public void IsActorExtensionIdentity_CategoryPresent_ReturnsTrue()
        {
            Assert.IsTrue(SidSnapHandler.IsActorExtensionIdentity(
                true, null, null));
        }
    }
}
