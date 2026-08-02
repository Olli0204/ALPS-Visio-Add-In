using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using VisioAddIn;
using MessageBox = System.Windows.MessageBox;

namespace VisioAddIn.Snapping
{
    /// <summary>
    /// manages the snapping of state extensions to states on sbd pages.
    /// </summary>
    public class SbdSnapHandler : SnapHandler
    {
        private readonly SBDPage foregroundPage;
        private SBDPage referencedBackgroundPage;

        private readonly ModelController modelController;

        public SbdSnapHandler(SBDPage foregroundPage, ModelController modelController) : base()
        {
            Debug.Print("Creating SbdSnapHandler for: " + foregroundPage.getNameU());
            this.modelController = modelController;
            this.foregroundPage = foregroundPage;

            referencedBackgroundPage = null;
        }

        /// <summary>
        /// checks for given snappingShape if it should snap
        /// shapes should snap when they are state extensions.
        /// </summary>
        /// <param name="shape">snappingShape to check</param>
        /// <returns>true if snappable, false otherwise</returns>
        protected override bool isShapeSnappable(IVShape shape)
        {
            Debug.Print("testing shape: " + shape.NameU + " - is snappable: "
                + shape.HasCategory(ALPSConstants.alpsShapeCategoryStateExtension)
                + " on: " + foregroundPage.getNameU()
                + " with background: "
                + (referencedBackgroundPage == null
                    ? "<none>"
                    : referencedBackgroundPage.getNameU()));
            return shape.HasCategory(ALPSConstants.alpsShapeCategoryStateExtension);
        }

        protected override void setBackPage(DiagramPage newProperty)
        {
            referencedBackgroundPage = newProperty as SBDPage;
        }
        
        /// <summary>
        /// A plug-in method which is called by the abstract base class 
        /// </summary>
        /// <param name="snappingShape"></param>
        protected override void handleDistantSnappedShapes(Shape snappingShape)
        {
            // Ask if the shapes should stay snapped
            WindowSnapMaintenance snapMain = new WindowSnapMaintenance(this, snappingShape, snappedShapes[snappingShape]);
            snapMain.ShowDialog();
        }

        protected override IEnumerable<Shape> getSnappableShapesOnBackgroundPage()
        {
            if (referencedBackgroundPage == null)
                return Enumerable.Empty<Shape>();

            SBDPageController referencedBackgroundPageController = modelController.getSbdPageController(referencedBackgroundPage);
            
            if (referencedBackgroundPageController == null) return new List<Shape>();
            return referencedBackgroundPageController.getPage().Shapes.Cast<Shape>()
                .Where(shape => shape.HasCategory(ALPSConstants.alpsShapeCategorySBDState)).ToList();
        }

        /// <summary>
        /// snaps the snappingShape to a another one, specified by name.
        /// </summary>
        /// <param name="snappingShape"></param>
        /// <param name="backgroundReferenceShapeName"></param>
        public override void snap(Shape snappingShape, string backgroundReferenceShapeName)
        {
            if (!isShapeSnappable(snappingShape)) return;
            backgroundReferenceShapeName =
                (backgroundReferenceShapeName ?? string.Empty)
                .Trim('\\', '"');
            if (string.IsNullOrWhiteSpace(backgroundReferenceShapeName))
            {
                unsnap(snappingShape);
                return;
            }

            if (snappedShapes.TryGetValue(
                snappingShape, out Shape currentReference)
                && MatchesReference(
                    currentReference, backgroundReferenceShapeName))
            {
                return;
            }

            Shape reference = getSnappableShapesOnBackgroundPage()
                .FirstOrDefault(shape => MatchesReference(
                    shape, backgroundReferenceShapeName));
            if (reference != null)
            {
                performSnap(snappingShape, reference);
                return;
            }

            unsnap(snappingShape);
            MessageBox.Show(
                string.Format(ALPSConstants.InputNotFound,
                    backgroundReferenceShapeName, snappingShape.NameU),
                "Error", MessageBoxButton.OK);
        }

        public void maintainSnap(Shape shape, Shape snapToShape)
        {
            adjustSize(shape, snapToShape);
        }

        /// <summary>
        /// unsnaps a snappingShape
        /// </summary>
        /// <param name="shape"></param>
        public override void unsnap(Shape shape)
        {
            if (!snappedShapes.ContainsKey(shape)) return;
            snappedShapes.Remove(shape);
            if (shape.CellExistsU[
                    ALPSConstants.cellPropertyCategoryPrefix + ALPSConstants.alpsPropertieTypeExtends +
                    ALPSConstants.cellValueSuffix, 0] == 0) return;
            Cell cell = shape.CellsU[ALPSConstants.cellValuePropertyExtends];
            cell.Formula = "";
        }

        /// <summary>
        /// called from SnapConfirmation.
        /// </summary>
        /// <param name="snap">true if it should snap, false if not</param>
        public override void performSnap(Shape snappingShape, Shape backgroundReferenceShape)
        {
            base.performSnap(snappingShape, backgroundReferenceShape);

            if (snappingShape.CellExistsU[ALPSConstants.cellValuePropertyExtends, 0] != 0)
            {
                Cell cell = snappingShape.CellsU[ALPSConstants.cellValuePropertyExtends];
                string snapToShapeId = backgroundReferenceShape.CellsU[ALPSConstants.cellValuePropertyModelComponentId].ResultStr[""];
                cell.Formula = "\"" + snapToShapeId + "\"";

            }
            if (snappingShape.CellExistsU[ALPSConstants.cellValuePropertyLabel, 0] != 0)
            {
                //Cell cell = snappingShape.CellsU[GlobalVariables.LableProp];
                //string snapToShapeLable = backgroundReferenceShape.CellsU["Prop." + ALPSConstants.alpsPropertieTypeLable + ".Value"].ResultStr[""];
                //cell.Formula = "\"" + GlobalVariables.LableExtension + snapToShapeLable + "\"";
            }
        }

        private static bool MatchesReference(
            Shape shape, string reference)
        {
            if (shape == null || string.IsNullOrWhiteSpace(reference))
                return false;

            string modelComponentId = null;
            if (shape.CellExistsU[
                ALPSConstants.cellValuePropertyModelComponentId, 0] != 0)
            {
                modelComponentId = shape.CellsU[
                    ALPSConstants.cellValuePropertyModelComponentId]
                    .ResultStr[""];
            }

            return string.Equals(modelComponentId, reference,
                    StringComparison.OrdinalIgnoreCase)
                || string.Equals(shape.Name, reference,
                    StringComparison.OrdinalIgnoreCase)
                || string.Equals(shape.NameU, reference,
                    StringComparison.OrdinalIgnoreCase);
        }

        

    }


}
