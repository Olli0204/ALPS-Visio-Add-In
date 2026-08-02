using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using VisioAddIn;

namespace VisioAddIn.Snapping
{
    /// <summary>
    /// Observes SID pages and checks for the snapping of subjects to subjects on the referenced background page
    /// </summary>
    public class SidSnapHandler : SnapHandler
    {
        /// <summary>
        /// The page which is currently observed and which contains a referenced background page
        /// </summary>
        private readonly SIDPage foregroundPage;

        /// <summary>
        /// The background to the currently active page
        /// </summary>
        private SIDPage referencedBackgroundPage;

        private readonly ModelController modelController;


        public SidSnapHandler(ModelController modelController, SIDPage foregroundPage) : base()
        {
            Debug.Print("Creating SidSnapHandler for: " + foregroundPage.getNameU());
            this.foregroundPage = foregroundPage;
            this.modelController = modelController;

            referencedBackgroundPage = null;
        }

        /// <summary>
        /// checks for given snappingShape if it should snap
        /// shapes should snap when they are actor extensions.
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
                    Debug.Print("Ignoring SID extension subshape: "
                        + shape.NameU);
                    return false;
                }
            }
            catch (COMException)
            {
                // Continue with identity checks for top-level RCWs that do
                // not expose ContainingShape reliably during a drop.
            }

            bool hasActorExtensionCategory = false;
            string masterName = null;
            string shapeName = null;

            try
            {
                hasActorExtensionCategory = shape.HasCategory(
                    ALPSConstants.alpsShapeCategoryActorExtension);
            }
            catch (COMException)
            {
                // Category rows can be initialized after ShapeAdded.
            }

            try
            {
                masterName = shape.Master?.NameU;
            }
            catch (COMException)
            {
                // Keep the independent name/category fallbacks available.
            }

            try
            {
                shapeName = shape.NameU;
            }
            catch (COMException)
            {
                // A later PinX/PinY event retries the same classification.
            }

            bool isSnappable = IsActorExtensionIdentity(
                hasActorExtensionCategory, masterName, shapeName);
            Debug.Print("SID snap candidate: "
                + (shapeName ?? "<unknown>")
                + "; master=" + (masterName ?? "<none>")
                + "; actorExtensionCategory="
                + hasActorExtensionCategory
                + "; snappable=" + isSnappable);
            return isSnappable;
        }

        internal static bool IsActorExtensionIdentity(
            bool hasActorExtensionCategory,
            string masterName, string shapeName)
        {
            return hasActorExtensionCategory
                || HasMasterIdentity(masterName,
                    ALPSConstants.alpsSIDMasterActorExtension)
                || HasMasterIdentity(masterName,
                    ALPSConstants.alpsSIDMasterActorGuardExtension)
                || HasMasterIdentity(masterName,
                    ALPSConstants.alpsSIDMasterActorMacroExtension)
                || HasMasterIdentity(shapeName,
                    ALPSConstants.alpsSIDMasterActorExtension)
                || HasMasterIdentity(shapeName,
                    ALPSConstants.alpsSIDMasterActorGuardExtension)
                || HasMasterIdentity(shapeName,
                    ALPSConstants.alpsSIDMasterActorMacroExtension);
        }

        private static bool HasMasterIdentity(
            string candidate, string expectedMasterName)
        {
            if (string.IsNullOrWhiteSpace(candidate)) return false;

            return string.Equals(candidate, expectedMasterName,
                    StringComparison.OrdinalIgnoreCase)
                || candidate.StartsWith(expectedMasterName + ".",
                    StringComparison.OrdinalIgnoreCase);
        }

        protected override void setBackPage(DiagramPage newProperty)
        {
            referencedBackgroundPage = newProperty as SIDPage;
        }

        /// <summary>
        /// snaps the snappingShape to a another one, specified by name.
        /// </summary>
        /// <param name="snappingShape">snappingShape to snap</param>
        /// <param name="backgroundReferenceShapeName">background snappingShape</param>
        public override void snap(Microsoft.Office.Interop.Visio.Shape snappingShape, string backgroundReferenceShapeName)
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

            // Do not retain a stale physical binding when a persisted
            // reference can no longer be resolved on the background page.
            unsnap(snappingShape);
        }

        /// <summary>
        /// A plug-in method which is called by the abstract base class 
        /// </summary>
        protected override void handleDistantSnappedShapes(Shape snappingShape)
        {
            WindowSnapMaintenance snapMaintenance =
                new WindowSnapMaintenance(
                    this, snappingShape,
                    snappedShapes[snappingShape]);
            snapMaintenance.ShowDialog();
        }

        /// <summary>
        /// unsnaps a given snappingShape and the associated sbd page.
        /// </summary>
        /// <param name="shape">snappingShape to unsnap</param>
        public override void unsnap(Shape shape)
        {
            if (!snappedShapes.TryGetValue(
                shape, out Shape referenceBackgroundShape))
            {
                return;
            }

            // Always remove the in-memory relation, even when an associated
            // SBD page or controller is missing during document teardown.
            snappedShapes.Remove(shape);
            SBDPage shapePage = null;
            SBDPage snapToShapePage = null;

            /*
            if (shape.CellExistsU[ALPSConstants.cellSubAdressHyperlinkLinkedSBD, 0] != 0)
            {
                shapePage = foregroundPage.getSbdPage(shape.CellsU[ALPSConstants.cellSubAdressHyperlinkLinkedSBD].Formula);
            }

            if (referenceBackgroundShape.CellExistsU[ALPSConstants.cellSubAdressHyperlinkLinkedSBD, 0] != 0)
            {
                snapToShapePage = referencedBackgroundPage.getSbdPage(referenceBackgroundShape.CellsU[ALPSConstants.cellSubAdressHyperlinkLinkedSBD].Formula);
            }*/

            shapePage = resolveLinkedSbdPage(foregroundPage, shape);
            snapToShapePage = resolveLinkedSbdPage(
                referencedBackgroundPage, referenceBackgroundShape);



            if (shapePage != null )
            {
                Debug.Print("setting to null");
                SBDPageController shapeController =
                    modelController.getSbdPageController(shapePage);
                shapeController?.setExtends(null);
            }
            if(snapToShapePage != null)
            {
                SBDPageController snapToController =
                    modelController.getSbdPageController(snapToShapePage);
                snapToController?.setNotExtended();
            }

            // Clear snappingShape contents that are related to snapping
            //if (shape.CellExistsU[ALPSConstants.cellSubAdressHyperlinkExtendedSubject, 0] != 0)
            //shape.CellsU[ALPSConstants.cellSubAdressHyperlinkExtendedSubject].Formula = "";
            if (shape.CellExistsU["Hyperlink." + ALPSConstants.alpsHyperlinksExtendedSubject, 0] != 0)
                shape.Hyperlinks.ItemU[ALPSConstants.alpsHyperlinksExtendedSubject].SubAddress = "";

            if (shape.CellExistsU[ALPSConstants.cellValuePropertyExtends, 0] != 0)
                shape.CellsU[ALPSConstants.cellValuePropertyExtends].Formula = "";

        }

        public void setModelUri(string newModelURI)
        {
            foreach (Shape shape in snappedShapes.Keys)
            {
                if (shape.CellExistsU[ALPSConstants.cellValuePropertyExtends, 0] == 0) continue;
                Cell cell = shape.CellsU[ALPSConstants.cellValuePropertyExtends];
                cell.Formula = "\"" + newModelURI + "#" + snappedShapes[shape].NameU + "\"";
            }
        }

        /// <summary>
        /// called after a snappingShape is added.
        /// if it's an actor extension, the diagram should be empty after adding 
        /// bc the new snappingShape isn't extending anything.
        /// </summary>
        /// <param name="shape"></param>
        internal void clearNewPage(Shape shape)
        {
            if (isShapeSnappable(shape))
            {
                SBDPage sbdPage = foregroundPage.getSbdPage(shape.NameU);
            }
        }

        /// <summary>
        /// called from SnapConfirmation.
        /// eventually snaps a snappingShape and the page associated with it.
        /// </summary>
        /// <param name="snap">true if it should snap, false if not</param>
        /// <param name="snappingShape"></param>
        /// <param name="backgroundReferenceShape"></param>
        public override void performSnap(Shape snappingShape, Shape backgroundReferenceShape)
        {
            if (referencedBackgroundPage == null) return;

            base.performSnap(snappingShape, backgroundReferenceShape);
            //Debug.Print("perform snap for: " + snappingShape.NameU + " and: " + backgroundReferenceShape.NameU);


            //if (snappingShape.CellExistsU[ALPSConstants.cellSubAdressHyperlinkExtendedSubject, 0] != 0)
            if (snappingShape.CellExistsU["Hyperlink." + ALPSConstants.alpsHyperlinksExtendedSubject, 0] != 0)
            {
                //Debug.Print("Writing cellSubAdressHyperlinkExtendedSubject");
                //Cell snappingShapeExtendedSubjectCell = snappingShape.CellsU[ALPSConstants.cellSubAdressHyperlinkExtendedSubject];
                //snappingShapeExtendedSubjectCell.Formula = "\"" + referencedBackgroundPage.getLayerForUser() + "/" + backgroundReferenceShape.NameU + "\"";
                snappingShape.Hyperlinks.ItemU[ALPSConstants.alpsHyperlinksExtendedSubject].SubAddress =  referencedBackgroundPage.getLayerForUser() + "/" + backgroundReferenceShape.NameU;
            }

            

            if (snappingShape.CellExistsU[ALPSConstants.cellValuePropertyExtends, 0] != 0)
            {
                //Debug.Print(" Writing Extends");
                Cell snappingShapeExtendsCell = snappingShape.CellsU[ALPSConstants.cellValuePropertyExtends];
                snappingShapeExtendsCell.Formula = "\"" + referencedBackgroundPage.getModelUriForUser() + "#" + backgroundReferenceShape.NameU + "\"";
            }
            /*
            if (snappingShape.CellExistsU[ALPSConstants.cellValuePropertyLabel, 0] != 0)
            {
                Cell cell = snappingShape.CellsU[GlobalVariables.LableProp];
                cell.Formula = "\"" + GlobalVariables.LableExtension + snapToShape.nameU + "\"";
            }*/

            SBDPage shapePage = null;
            SBDPage snapToShapePage = null;

            //Debug.Print("shape: " + snappingShape.NameU + " Link cell: " + (snappingShape.CellExistsU["Hyperlink.linkedSBD", 0] != 0) + " SubLink cell: " + (snappingShape.CellExistsU["Hyperlink.linkedSBD.SubAddress", 0] != 0));
           // Debug.Print(" - ALPSConstants.cellSubAdressHyperlinkLinkedSBD: " + ALPSConstants.cellSubAdressHyperlinkLinkedSBD);
            shapePage = resolveLinkedSbdPage(
                foregroundPage, snappingShape);
            snapToShapePage = resolveLinkedSbdPage(
                referencedBackgroundPage, backgroundReferenceShape);
            //Debug.Print("foregroundPage: " + foregroundPage.getNameU() + " referencedBackgroundPage: " + referencedBackgroundPage.getNameU());
           //Debug.Print("shapePage: " + shapePage.getNameU() + " - snapToShapePage: " + snapToShapePage.getNameU());


            if (shapePage == null || snapToShapePage == null)
            {
                Debug.Print("SID snap could not establish the SBD relation: "
                    + "foreground=" + foregroundPage.getNameU()
                    + ", extension=" + snappingShape.NameU
                    + ", extensionSBD="
                    + (shapePage == null ? "<none>" : shapePage.getNameU())
                    + ", background=" + referencedBackgroundPage.getNameU()
                    + ", subject=" + backgroundReferenceShape.NameU
                    + ", subjectSBD="
                    + (snapToShapePage == null
                        ? "<none>"
                        : snapToShapePage.getNameU()));
                return;
            }

            SBDPageController shapePageC = modelController.getSbdPageController(shapePage);
            SBDPageController snapToShapePageC = modelController.getSbdPageController(snapToShapePage);

            // Gets the shapeType of the snappingShape currently snapping
            SBDPage oldExtends = shapePage.getExtends();
            string shapeType = snappingShape.CellsU[ALPSConstants.cellValuePropertyModelComponentType].Formula;
            shapeType = shapeType.Replace("\"", "");

            // Do not set Extends for SBD if it is a makro extension
            if (!snappingShape.HasCategory(ALPSConstants.MacroExtension))
            {
                if (snapToShapePageC == null || shapePageC == null)
                    return;

                snapToShapePageC.setExtended(shapePage);
                //Debug.Print("extension set");
                shapePageC.setExtends(snapToShapePage);
                Debug.Print("Established SBD extension relation: "
                    + shapePage.getNameU() + " -> "
                    + snapToShapePage.getNameU());
            }

            if (oldExtends == null) return;
            SBDPageController oldExtendsC = modelController.getSbdPageController(oldExtends);
            oldExtendsC?.setNotExtended();

        }

        /// <summary>
        /// checks for a given page if there are standard actors shapes should be snapping to.
        /// </summary>
        /// <param name="backPage"></param>
        /// <returns></returns>
        protected override IEnumerable<Shape> getSnappableShapesOnBackgroundPage()
        {
            if (referencedBackgroundPage == null)
                return Enumerable.Empty<Shape>();

            SIDPageController referencedBackgroundPageController = modelController.getSidPageController(referencedBackgroundPage);
            return referencedBackgroundPageController == null ? new List<Shape>() :
                referencedBackgroundPageController.getPage().Shapes.Cast<Shape>()
                    .Where(shape => shape.HasCategory(ALPSConstants.alpsShapeCategoryStandardActor)).ToList();
        }

        private static bool MatchesReference(
            Shape shape, string reference)
        {
            if (shape == null || string.IsNullOrWhiteSpace(reference))
                return false;

            return string.Equals(
                    shape.Name, reference,
                    StringComparison.OrdinalIgnoreCase)
                || string.Equals(
                    shape.NameU, reference,
                    StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Resolves the behavior page linked from a subject. Older stencil
        /// macros can retain the SID number that existed while a compound
        /// GuardExtension was assembled (for example SID_1 in the page name
        /// while the final shape lives on SID_5). The stable suffix beginning
        /// with the master identity still identifies the behavior uniquely.
        /// </summary>
        private SBDPage resolveLinkedSbdPage(
            SIDPage sidPage, Shape subjectShape)
        {
            if (sidPage == null || subjectShape == null) return null;

            string linkedPageReference = null;
            try
            {
                if (subjectShape.CellExistsU[
                        "Hyperlink."
                        + ALPSConstants.alpsHyperlinkTypeLinkedSBD,
                        0] != 0)
                {
                    linkedPageReference = subjectShape.Hyperlinks.ItemU[
                        ALPSConstants.alpsHyperlinkTypeLinkedSBD]
                        .SubAddress;
                }
            }
            catch (COMException)
            {
                // Continue with the identity-based fallbacks below.
            }

            SBDPage exact = sidPage.getSbdPage(linkedPageReference);
            if (exact != null) return exact;

            SBDPage registered =
                modelController.ensureLinkedSbdPageRegistered(
                    sidPage, linkedPageReference);
            if (registered != null) return registered;

            string shapeName = null;
            string masterName = null;
            try
            {
                shapeName = subjectShape.NameU;
                masterName = subjectShape.Master?.NameU;
            }
            catch (COMException)
            {
                // A unique page is still a safe final fallback.
            }

            // FullySpecifiedSubject behavior pages conventionally use the
            // subject shape's NameU. This fallback also covers legacy stencil
            // instances whose linkedSBD hyperlink is empty or initialized
            // after the first controller scan.
            registered = modelController.ensureLinkedSbdPageRegistered(
                sidPage, shapeName);
            if (registered != null) return registered;

            string identitySuffix = getMasterIdentitySuffix(
                shapeName, masterName);
            if (!string.IsNullOrWhiteSpace(identitySuffix))
            {
                List<SBDPage> identityMatches = sidPage.getSbdPages()
                    .Where(page => page.getNameU().IndexOf(
                        identitySuffix,
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
                if (identityMatches.Count == 1)
                {
                    Debug.Print("Resolved linked SBD by subject identity '"
                        + identitySuffix + "': "
                        + identityMatches[0].getNameU());
                    return identityMatches[0];
                }
            }

            IList<SBDPage> availablePages = sidPage.getSbdPages();
            if (availablePages.Count == 1)
            {
                Debug.Print("Resolved the only SBD on "
                    + sidPage.getNameU() + " for "
                    + (shapeName ?? "<unknown subject>") + ": "
                    + availablePages[0].getNameU());
                return availablePages[0];
            }

            return null;
        }

        internal static string getMasterIdentitySuffix(
            string shapeName, string masterName)
        {
            if (string.IsNullOrWhiteSpace(shapeName)
                || string.IsNullOrWhiteSpace(masterName))
            {
                return null;
            }

            int masterStart = shapeName.IndexOf(
                masterName, StringComparison.OrdinalIgnoreCase);
            return masterStart < 0
                ? null
                : shapeName.Substring(masterStart);
        }
    }
}
