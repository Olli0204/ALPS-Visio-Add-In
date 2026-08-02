using ALPS_Visio_AddIn_rewrite.BpmnConversion;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace ALPS_Visio_AddIn_rewrite.Tests.BpmnConversion
{
    [TestClass]
    public sealed class BpmnConversionTests
    {
        [TestMethod]
        [TestCategory("Integration")]
        [Timeout(60000)]
        public void Convert_VacationRequestFixture_WritesValidBpmnDiagram()
        {
            string output = TestResources.TemporaryFile(".bpmn");
            try
            {
                BpmnConversionController.Convert(
                    TestResources.Fixture("[Test]_Vacation_Request.owl"),
                    output);

                XDocument document = XDocument.Load(output);
                XNamespace bpmn =
                    "http://www.omg.org/spec/BPMN/20100524/MODEL";
                XNamespace bpmnDi =
                    "http://www.omg.org/spec/BPMN/20100524/DI";
                XNamespace di =
                    "http://www.omg.org/spec/DD/20100524/DI";
                Assert.AreEqual(
                    bpmn + "definitions", document.Root.Name);
                Assert.IsTrue(
                    document.Descendants(bpmn + "participant").Count() >= 2);
                Assert.IsTrue(
                    document.Descendants(bpmnDi + "BPMNShape").Any());
                Assert.IsTrue(document
                    .Descendants(bpmnDi + "BPMNEdge")
                    .All(edge => edge.Elements(di + "waypoint").Count() >= 2));
            }
            finally
            {
                if (File.Exists(output))
                    File.Delete(output);
            }
        }

        [TestMethod]
        public void ValidateSerializedBpmn_DuplicateIds_ThrowsInvalidData()
        {
            string output = TestResources.TemporaryFile(".bpmn");
            try
            {
                File.WriteAllText(
                    output,
                    "<definitions xmlns=\"http://www.omg.org/spec/BPMN/20100524/MODEL\">"
                    + "<process id=\"process\"><task id=\"duplicate\" />"
                    + "<task id=\"duplicate\" /></process></definitions>");

                InvalidDataException exception =
                    Assert.ThrowsExactly<InvalidDataException>(() =>
                        BpmnConversionController.ValidateSerializedBpmn(output));

                StringAssert.Contains(exception.Message, "mehrfach");
            }
            finally
            {
                if (File.Exists(output))
                    File.Delete(output);
            }
        }
    }
}
