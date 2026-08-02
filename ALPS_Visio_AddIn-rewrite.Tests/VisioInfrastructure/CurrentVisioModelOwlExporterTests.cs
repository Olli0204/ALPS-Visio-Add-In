using ALPS_Visio_AddIn_rewrite.VisioInfrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace ALPS_Visio_AddIn_rewrite.Tests.VisioInfrastructure
{
    [TestClass]
    public sealed class CurrentVisioModelOwlExporterTests
    {
        [TestMethod]
        public void MacroCommands_SidStencilV101_MatchExportContract()
        {
            Assert.AreEqual(
                "ALPS_RDFOWLExporter.createProcessRDFOWL",
                CurrentVisioModelOwlExporter.ExportMacro);
            Assert.AreEqual(
                "ALPS_RDFOWLExporter."
                + "includeMinimal2DVisualisationDataInOWL = True",
                CurrentVisioModelOwlExporter.EnableMinimal2DExport);
        }

        [TestMethod]
        public void GetOutputFilePath_SavedVsdx_MatchesStencilMacro()
        {
            string directory = Path.Combine(
                Path.GetTempPath(), "ALPS Models");

            string result = CurrentVisioModelOwlExporter
                .GetOutputFilePath(
                    directory, "Vacation Request.vsdx");

            Assert.AreEqual(
                Path.Combine(directory, "Vacation_Request.owl"),
                result);
        }

        [TestMethod]
        public void GetOutputFilePath_UnsavedDrawing_RejectsExport()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
                CurrentVisioModelOwlExporter.GetOutputFilePath(
                    string.Empty, "Drawing1.vsdx"));
        }

        [TestMethod]
        public void GetOutputFilePath_WebOnlyDrawing_RejectsHandoff()
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
                CurrentVisioModelOwlExporter.GetOutputFilePath(
                    "https://example.sharepoint.com/models/",
                    "Vacation Request.vsdx"));
        }
    }
}
