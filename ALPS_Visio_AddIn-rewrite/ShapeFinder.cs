using ALPS_Visio_AddIn_rewrite.VisioInfrastructure;

namespace ALPS_Visio_AddIn_rewrite
{
    /// <summary>
    /// Source-compatible facade for locating installed ALPS/PASS stencil files.
    /// </summary>
    public static class ShapeFinder
    {
        private static readonly StencilFileLocator Locator =
            new StencilFileLocator(() => Globals.ThisAddIn.Application.MyShapesPath);

        /// <summary>
        /// Returns the name of the newest SID stencil in the Visio My Shapes folder.
        /// </summary>
        public static string getSIDName()
        {
            return Locator.GetSidName();
        }

        /// <summary>
        /// Returns the name of the newest SBD stencil in the Visio My Shapes folder.
        /// </summary>
        public static string getSBDName()
        {
            return Locator.GetSbdName();
        }
    }
}
