using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAddIn.Snapping;

namespace ALPS_Visio_AddIn_rewrite.Tests.Snapping
{
    [TestClass]
    public sealed class ModelControllerTests
    {
        [TestMethod]
        public void MatchesSidPageReference_UniversalPageName_ReturnsTrue()
        {
            Assert.IsTrue(ModelController.MatchesSidPageReference(
                "SID_1", "Base Layer", "SID_1"));
        }

        [TestMethod]
        public void MatchesSidPageReference_QuotedLayerName_ReturnsTrue()
        {
            Assert.IsTrue(ModelController.MatchesSidPageReference(
                "\"Base Layer\"", "Base Layer", "SID_1"));
        }

        [TestMethod]
        public void MatchesSidPageReference_UnrelatedName_ReturnsFalse()
        {
            Assert.IsFalse(ModelController.MatchesSidPageReference(
                "SID_7", "Base Layer", "SID_1"));
        }
    }
}
