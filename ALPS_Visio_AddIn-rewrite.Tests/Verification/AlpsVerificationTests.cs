using ALPS_Visio_AddIn_rewrite.Verification;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;

namespace ALPS_Visio_AddIn_rewrite.Tests.Verification
{
    [TestClass]
    public sealed class AlpsVerificationTests
    {
        [TestMethod]
        public void LocalName_UriAndEncodedValue_ReturnsDecodedLocalName()
        {
            Assert.AreEqual(
                "Subject A",
                OwlVerificationModel.LocalName(
                    "<https://example.org/model#Subject%20A>"));
            Assert.AreEqual(
                "SubjectB",
                OwlVerificationModel.LocalName(
                    "https://example.org/model/SubjectB"));
            Assert.AreEqual(string.Empty, OwlVerificationModel.LocalName(null));
        }

        [TestMethod]
        public void ReferenceMatches_ComponentIdOrUriLocalName_ReturnsTrue()
        {
            OwlVerificationElement element = new OwlVerificationElement(
                "https://example.org/model#SubjectA")
            {
                ComponentId = "subject-id"
            };

            Assert.IsTrue(OwlVerificationModel.ReferenceMatches(
                "subject-id", element));
            Assert.IsTrue(OwlVerificationModel.ReferenceMatches(
                "https://other.example/#SubjectA", element));
            Assert.IsFalse(OwlVerificationModel.ReferenceMatches(
                "SubjectB", element));
        }

        [TestMethod]
        public void Load_SpecificationFixture_ExtractsSubjectsAndRestriction()
        {
            OwlVerificationModel model = OwlVerificationModel.Load(
                SpecificationFixture);

            Assert.AreEqual(
                2,
                model.ElementsOfKind(OwlElementKind.Subject).Count());
            Assert.AreEqual(
                1,
                model.ElementsOfKind(
                    OwlElementKind.CommunicationRestriction).Count());
            Assert.IsNotNull(model.Resolve("RestrictedApprover"));
        }

        [TestMethod]
        public void Verify_RegressionFixtures_ReportsForbiddenCommunication()
        {
            AlpsVerificationService service =
                new AlpsVerificationService();

            VerificationReport report = service.Verify(
                SpecificationFixture, ImplementationFixture);

            Assert.AreEqual(1, report.ErrorCount);
            Assert.AreEqual(0, report.WarningCount);
            Assert.IsFalse(report.PassedSupportedChecks);
            VerificationFinding finding = report.Findings.Single(candidate =>
                candidate.Code == "RESTRICTION-003");
            StringAssert.Contains(
                finding.ImplementationElement, "ForbiddenRequest");
            Assert.IsTrue(report.Findings.Any(candidate =>
                candidate.Code == "SCOPE-001"
                && candidate.Severity == VerificationSeverity.Information));
        }

        [TestMethod]
        public void Verify_RestrictionEndpointsReversed_StillReportsViolation()
        {
            string reversedImplementation =
                TestResources.TemporaryFile(".owl");
            try
            {
                string owl = File.ReadAllText(ImplementationFixture)
                    .Replace(
                        "hasSender rdf:resource=\"https://example.org/alps-verification/implementation#ConcreteRequester\"",
                        "hasSender rdf:resource=\"https://example.org/alps-verification/implementation#ConcreteApprover\"")
                    .Replace(
                        "hasReceiver rdf:resource=\"https://example.org/alps-verification/implementation#ConcreteApprover\"",
                        "hasReceiver rdf:resource=\"https://example.org/alps-verification/implementation#ConcreteRequester\"");
                File.WriteAllText(reversedImplementation, owl);

                VerificationReport report =
                    new AlpsVerificationService().Verify(
                        SpecificationFixture, reversedImplementation);

                Assert.IsTrue(report.Findings.Any(candidate =>
                    candidate.Code == "RESTRICTION-003"));
            }
            finally
            {
                if (File.Exists(reversedImplementation))
                    File.Delete(reversedImplementation);
            }
        }

        [TestMethod]
        public void Load_MissingFile_ThrowsFileNotFound()
        {
            string missing = TestResources.TemporaryFile(".owl");

            Assert.ThrowsExactly<FileNotFoundException>(
                () => OwlVerificationModel.Load(missing));
        }

        [TestMethod]
        public void VerificationReport_FindingsAdded_ComputesSummary()
        {
            VerificationReport report =
                new VerificationReport("spec.owl", "impl.owl");
            report.Findings.Add(new VerificationFinding
            {
                Severity = VerificationSeverity.Error
            });
            report.Findings.Add(new VerificationFinding
            {
                Severity = VerificationSeverity.Warning
            });

            Assert.AreEqual(1, report.ErrorCount);
            Assert.AreEqual(1, report.WarningCount);
            Assert.IsFalse(report.PassedSupportedChecks);
        }

        private static string SpecificationFixture =>
            TestResources.Fixture(
                "[Test]_ALPS_Verification_Specification.owl");

        private static string ImplementationFixture =>
            TestResources.Fixture(
                "[Test]_ALPS_Verification_Implementation.owl");
    }
}
