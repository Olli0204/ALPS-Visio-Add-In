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

        [TestMethod]
        public void IsStateReferenceIdentity_StateExtensionMaster_ReturnsTrue()
        {
            Assert.IsTrue(SbdSnapHandler.IsStateReferenceIdentity(
                false, "StateExtension", null, null));
        }

        [TestMethod]
        public void IsStateReferenceIdentity_StateExtensionInstance_ReturnsTrue()
        {
            Assert.IsTrue(SbdSnapHandler.IsStateReferenceIdentity(
                false, null, "StateExtension.42", null));
        }

        [TestMethod]
        public void IsStateReferenceIdentity_ComponentType_ReturnsTrue()
        {
            Assert.IsTrue(SbdSnapHandler.IsStateReferenceIdentity(
                false, null, null, "StateReference"));
        }

        [TestMethod]
        public void IsStateReferenceIdentity_CategoryPresent_ReturnsTrue()
        {
            Assert.IsTrue(SbdSnapHandler.IsStateReferenceIdentity(
                true, null, null, null));
        }

        [TestMethod]
        public void IsStateReferenceIdentity_NormalState_ReturnsFalse()
        {
            Assert.IsFalse(SbdSnapHandler.IsStateReferenceIdentity(
                false, "FunctionState", "FunctionState.42",
                "DoState"));
        }

        [TestMethod]
        public void GetMasterIdentitySuffix_GuardExtensionWithChangedSid_ReturnsStableSuffix()
        {
            Assert.AreEqual("GuardExtension_3",
                SidSnapHandler.getMasterIdentitySuffix(
                    "SID_5_GuardExtension_3", "GuardExtension"));
        }

        [TestMethod]
        public void GetMasterIdentitySuffix_MasterNotInShape_ReturnsNull()
        {
            Assert.IsNull(SidSnapHandler.getMasterIdentitySuffix(
                "SID_5_ActorExtension_3", "GuardExtension"));
        }
    }
}
