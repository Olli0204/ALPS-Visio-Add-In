using Microsoft.Office.Interop.Visio;
using System.Linq;
using VisioAddIn;

namespace VisioAddIn.Snapping
{
    /// <summary>
    /// Recognizes fully initialized SID and SBD pages from their ShapeSheet cells.
    /// </summary>
    internal static class DiagramPageClassifier
    {
        private static readonly string[] RequiredSidCells =
        {
            ALPSConstants.cellValuePropertyPageModelURI,
            ALPSConstants.cellValuePropertyPageType,
            ALPSConstants.cellValuePropertyPageLayer,
            ALPSConstants.cellValuePropertyPageModelVersion,
            ALPSConstants.cellValuePropertyPriorityOrderNumber
        };

        public static bool IsSid(Page page)
        {
            if (RequiredSidCells.Any(
                cell => page.PageSheet.CellExistsU[cell, 1] == 0))
            {
                return false;
            }

            string pageType =
                page.PageSheet.CellsU[ALPSConstants.cellValuePropertyPageType].Formula;
            return pageType.Contains("SubjectInteraction");
        }

        public static bool IsSbd(Page page)
        {
            return page.PageSheet.CellExistsU[
                ALPSConstants.cellValuePropertySBDLinkedSubjectID, 1] != 0;
        }
    }
}
