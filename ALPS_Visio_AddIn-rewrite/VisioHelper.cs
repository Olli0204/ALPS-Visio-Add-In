using alps.net.api.ALPS;
using alps.net.api.StandardPASS;
using System;
using System.Collections.Generic;
using ALPS_Visio_AddIn_rewrite.OWLShapes.Layout;
using ALPS_Visio_AddIn_rewrite.VisioInfrastructure;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite
{
    public static class VisioHelper
    {
        private static readonly VisioStencilRepository StencilRepository =
            new VisioStencilRepository(
                () => Globals.ThisAddIn.Application.Documents,
                message => System.Windows.Forms.MessageBox.Show(message));

        private static readonly VisioStencilSession StencilSession =
            new VisioStencilSession(
                () => Globals.ThisAddIn.Application,
                () => Globals.ThisAddIn.GetDrawingDocument(),
                StencilRepository);

        private static readonly VisioPageFactory PageFactory =
            new VisioPageFactory(() => Globals.ThisAddIn.Application);

        public static void setVBAListenersRunning(Boolean newStatus)
        {
            Visio.IVDocument myActiveDocument =
                Globals.ThisAddIn.GetDrawingDocument();

            if (myActiveDocument == null) return;

            SetBool(myActiveDocument.DocumentSheet, Constants.Properties.InteropWithVSTOShouldListenersRun, newStatus);
        }

        public enum VisioStencils
        {
            SID_STENCIL,
            SBD_STENCIL
        }

        /// <summary>
        /// Opens the latest SID-Stencil file from specified shape-folder
        /// </summary>
        /// <returns>The specified stencil file or null</returns>
        public static Visio.Document openStencil(VisioStencils stencil)
        {
            return StencilRepository.Open(stencil);
        }

        /// <summary>
        /// Opens both master sources for automated import without executing VBA.
        /// </summary>
        public static void OpenImportStencils()
        {
            StencilSession.OpenImportStencils();
        }

        /// <summary>
        /// Opens both stencils for interactive use, allowing their VBA macros.
        /// </summary>
        public static void OpenInteractiveStencils()
        {
            StencilSession.OpenInteractiveStencils();
        }

        public enum ShapeType
        {
            SBD, SID
        }

        public enum GraphLayoutDirection
        {
            TopDown = 1,
            LeftRight = 2
        }

        public static void AutoArrangePage(Visio.IVPage page, GraphLayoutDirection direction)
        {
            VisioAutoArrange.Arrange(page, direction);
        }

        public static Visio.Shape Place(string shapeType, Visio.Page page)
        {
            return StencilRepository.Place(shapeType, page);
        }

        public static VisioStencils GetStencil(string shapeType)
        {
            return StencilRepository.GetStencil(shapeType);
        }

        public static void SetHyperlink(Visio.Shape shape, string property, string value)
        {
            VisioShapeSheet.SetHyperlink(shape, property, value);
        }

        public static void SetProperty(Visio.Shape shape, string property, string value)
        {
            VisioShapeSheet.SetProperty(shape, property, value);
        }

        public static void SetBool(Visio.Shape shape, string property, bool value)
        {
            VisioShapeSheet.SetBooleanProperty(shape, property, value);
        }

        public static void SetSize(Visio.Shape shape, string cell, double value)
        {
            VisioShapeSheet.SetNumber(shape, cell, value);
        }

        public static double GetSize(Visio.Shape shape, string cell)
        {
            return VisioShapeSheet.GetNumber(shape, cell);
        }

        public static void SetPropertyU(Visio.Shape shape, string property, object value)
        {
            VisioShapeSheet.SetUniversalProperty(shape, property, value);
        }

        public static void SetPropertyULiteral(Visio.Shape shape, string property, object value)
        {
            VisioShapeSheet.SetUniversalLiteralProperty(shape, property, value);
        }

        public static void SetPropertyFormulaU(Visio.Shape shape, string property, object value)
        {
            VisioShapeSheet.SetUniversalFormulaProperty(shape, property, value);
        }

        public static void SetUser(Visio.Shape shape, string user, object value)
        {
            VisioShapeSheet.SetUserCell(shape, user, value);
        }

        public static void SetSizeMM(Visio.Shape shape, string cell, object value)
        {
            VisioShapeSheet.SetMillimeters(shape, cell, value);
        }

        public static void ConfigureFallbackSbdRouting(Visio.Page page)
        {
            VisioRouting.ConfigureFallbackSbd(page);
        }

        public static void ConfigureFallbackTransitionRouting(Visio.Shape connector, bool isFeedback)
        {
            VisioRouting.ConfigureFallbackTransition(connector, isFeedback);
        }

        public static void FinalizeFallbackTransitionRouting(Visio.Shape connector, bool isFeedback)
        {
            VisioRouting.FinalizeFallbackTransition(connector, isFeedback);
        }

        public static void EnsureImportedDirectionalLine(Visio.Shape connector)
        {
            VisioRouting.EnsureImportedDirectionalLine(connector);
        }

        public static Visio.Page CreateSIDPage(string name, string nameU, string modelURI, string extends, string implements, string priority)
        {
            return PageFactory.CreateSidPage(name, nameU, modelURI, extends, implements, priority);
        }

        public static Visio.Page CreateSBDPage(Visio.Page sidPage, string name, string nameU, Visio.Shape subjectShape)
        {
            return PageFactory.CreateSbdPage(sidPage, name, nameU, subjectShape);
        }

        public static List<ISimple2DVisualizationPoint> GetBounds(PASSProcessModelElement element)
        {
            return VisioBoundsProvider.GetBounds(element);
        }
    }
}
