using ALPS_Visio_AddIn_rewrite.VisioInfrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ALPS_Visio_AddIn_rewrite.Tests.VisioInfrastructure
{
    [TestClass]
    public class VisioRoutingTests
    {
        [DataTestMethod]
        [DataRow(0d, "0", true)]
        [DataRow(1d, "1", false)]
        [DataRow(23d, "23", false)]
        [DataRow(24d, "24", true)]
        [DataRow(254d, "USE(\"PASS Transition\")", true)]
        public void RequiresPatternFallback_LinePatternValue_ReturnsExpected(
            double value, string formula, bool expected)
        {
            Assert.AreEqual(expected,
                VisioRouting.RequiresPatternFallback(value, formula));
        }

        [DataTestMethod]
        [DataRow(0d, "0", true)]
        [DataRow(1d, "1", false)]
        [DataRow(45d, "45", false)]
        [DataRow(46d, "46", true)]
        [DataRow(254d, "use(\"PASS Arrow\")", true)]
        public void RequiresArrowFallback_ArrowValue_ReturnsExpected(
            double value, string formula, bool expected)
        {
            Assert.AreEqual(expected,
                VisioRouting.RequiresArrowFallback(value, formula));
        }
    }
}
