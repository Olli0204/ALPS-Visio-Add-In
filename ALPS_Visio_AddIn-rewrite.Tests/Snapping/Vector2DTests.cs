using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAddIn.util;

namespace ALPS_Visio_AddIn_rewrite.Tests.Snapping
{
    [TestClass]
    public sealed class Vector2DTests
    {
        [TestMethod]
        public void Constructor_CoordinatesProvided_ReturnsCoordinates()
        {
            Vector2D vector = new Vector2D(12.5, -4.25);

            Assert.AreEqual(12.5, vector.getX());
            Assert.AreEqual(-4.25, vector.getY());
        }

        [TestMethod]
        public void IsNearTo_DifferenceBelowThreshold_ReturnsTrue()
        {
            Vector2D origin = new Vector2D(10, 10);
            Vector2D nearby = new Vector2D(29.999, -9.999);

            Assert.IsTrue(origin.isNearTo(nearby));
        }

        [TestMethod]
        public void IsNearTo_DifferenceEqualsThreshold_ReturnsFalse()
        {
            Vector2D origin = new Vector2D(10, 10);

            Assert.IsFalse(origin.isNearTo(new Vector2D(30, 10)));
            Assert.IsFalse(origin.isNearTo(new Vector2D(10, -10)));
        }

        [TestMethod]
        public void IsNearTo_OnlyOneAxisIsOutsideThreshold_ReturnsFalse()
        {
            Vector2D origin = new Vector2D(0, 0);

            Assert.IsFalse(origin.isNearTo(new Vector2D(5, 25)));
        }
    }
}
