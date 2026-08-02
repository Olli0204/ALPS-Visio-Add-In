using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
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
            if (shape == null) return false;

            try
            {
                Shape containingShape = shape.ContainingShape;
                if (containingShape != null
                    && containingShape.Type
                    == (short)VisShapeTypes.visTypeGroup)
                {
                    Debug.Print("Ignoring StateReference subshape: "
                        + shape.NameU);
                    return false;
                }
            }
            catch (COMException)
            {
                // Continue with identity checks for top-level RCWs that do
                // not expose ContainingShape reliably during a drop.
            }

            bool hasStateExtensionCategory = false;
            string masterName = null;
            string shapeName = null;
            string componentType = null;

            try
            {
                hasStateExtensionCategory = shape.HasCategory(
                    ALPSConstants.alpsShapeCategoryStateExtension);
            }
            catch (COMException)
            {
                // Category rows may still be initialized by EventDrop.
            }

            try
            {
                masterName = shape.Master?.NameU;
            }
            catch (COMException)
            {
                // Keep the independent identity fallbacks available.
            }

            try
            {
                shapeName = shape.NameU;
            }
            catch (COMException)
            {
                // A later movement event can retry classification.
            }

            try
            {
                if (shape.CellExistsU[
                        ALPSConstants.cellValuePropertyModelComponentType,
                        0] != 0)
                {
                    componentType = shape.CellsU[
                        ALPSConstants.cellValuePropertyModelComponentType]
                        .ResultStr[""];
                }
            }
            catch (COMException)
            {
                // The property row is another optional fallback.
            }

            bool isSnappable = IsStateReferenceIdentity(
                hasStateExtensionCategory, masterName, shapeName,
                componentType);
            Debug.Print("SBD snap candidate: "
                + (shapeName ?? "<unknown>")
                + "; master=" + (masterName ?? "<none>")
                + "; componentType=" + (componentType ?? "<none>")
                + "; stateExtensionCategory="
                + hasStateExtensionCategory
                + "; page=" + foregroundPage.getNameU()
                + "; background="
                + (referencedBackgroundPage == null
                    ? "<none>"
                    : referencedBackgroundPage.getNameU())
                + "; snappable=" + isSnappable);
            return isSnappable;
        }

        internal static bool IsStateReferenceIdentity(
            bool hasStateExtensionCategory,
            string masterName, string shapeName, string componentType)
        {
            return hasStateExtensionCategory
                || HasIdentity(masterName,
                    ALPSConstants.alpsSBDMasterStateExtension)
                || HasIdentity(shapeName,
                    ALPSConstants.alpsSBDMasterStateExtension)
                || string.Equals(componentType, "StateReference",
                    StringComparison.OrdinalIgnoreCase)
                || string.Equals(componentType, "StateExtension",
                    StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasIdentity(
            string candidate, string expectedIdentity)
        {
            if (string.IsNullOrWhiteSpace(candidate)) return false;

            return string.Equals(candidate, expectedIdentity,
                    StringComparison.OrdinalIgnoreCase)
                || candidate.StartsWith(expectedIdentity + ".",
                    StringComparison.OrdinalIgnoreCase);
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
