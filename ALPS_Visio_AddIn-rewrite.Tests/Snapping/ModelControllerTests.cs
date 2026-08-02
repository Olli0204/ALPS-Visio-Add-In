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

        [TestMethod]
        public void MatchesDocumentSidPageReference_EncodedPageId_ReturnsTrue()
        {
            Assert.IsTrue(ModelController.MatchesDocumentSidPageReference(
                "SID_1", "Base Layer", "LayerPage", "Base", 1));
        }

        [TestMethod]
        public void MatchesDocumentSidPageReference_DisplayName_ReturnsTrue()
        {
            Assert.IsTrue(ModelController.MatchesDocumentSidPageReference(
                "Base SID", "Layer A", "SID_9", "Base SID", 9));
        }
    }
}
