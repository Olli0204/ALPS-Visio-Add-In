using ALPS_Visio_AddIn_rewrite.NlpChecking;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ALPS_Visio_AddIn_rewrite.Tests.NlpChecking
{
    [TestClass]
    public sealed class NlpNameClassifierTests
    {
        [TestMethod]
        public void EnsureTrained_BundledDataset_ReturnsExpectedExampleCount()
        {
            NlpNameClassifier classifier = new NlpNameClassifier();

            int count = classifier.EnsureTrained();

            Assert.AreEqual(680, count);
            Assert.AreEqual(count, classifier.EnsureTrained());
        }

        [TestMethod]
        public void Predict_KnownExamples_DistinguishesGoodAndGenericLabels()
        {
            NlpNameClassifier classifier = new NlpNameClassifier();

            NlpNamePrediction valid = classifier.Predict(
                "Customer", "FullySpecifiedSubject");
            NlpNamePrediction invalid = classifier.Predict(
                "Person", "FullySpecifiedSubject");

            Assert.IsTrue(valid.IsValid);
            Assert.IsFalse(invalid.IsValid);
            Assert.IsTrue(valid.Probability >= 0.0f
                && valid.Probability <= 1.0f);
            Assert.IsTrue(invalid.Probability >= 0.0f
                && invalid.Probability <= 1.0f);
        }

        [TestMethod]
        public void Predict_UnknownShapeType_ReturnsBoundedFallbackPrediction()
        {
            NlpNamePrediction prediction = new NlpNameClassifier().Predict(
                "Review request", "CustomShape");

            Assert.IsTrue(prediction.Probability >= 0.0f);
            Assert.IsTrue(prediction.Probability <= 1.0f);
        }

        [TestMethod]
        public void Retrain_SameInput_ReturnsDeterministicPrediction()
        {
            NlpNameClassifier classifier = new NlpNameClassifier();
            NlpNamePrediction before = classifier.Predict(
                "Send approval", "SendState");

            int count = classifier.Retrain();
            NlpNamePrediction after = classifier.Predict(
                "Send approval", "SendState");

            Assert.AreEqual(680, count);
            Assert.AreEqual(before.IsValid, after.IsValid);
            Assert.AreEqual(before.Probability, after.Probability, 0.000001f);
        }
    }
}
