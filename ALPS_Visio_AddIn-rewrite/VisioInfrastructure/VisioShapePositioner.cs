using alps.net.api.ALPS;
using System;
using System.Collections.Generic;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.VisioInfrastructure
{
    /// <summary>
    /// Applies normalized ALPS visualization bounds to a Visio shape.
    /// </summary>
    internal static class VisioShapePositioner
    {
        public static void Apply(Visio.Shape shape, Visio.Page page,
            IList<ISimple2DVisualizationPoint> bounds)
        {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (bounds == null || bounds.Count < 2) return;

            double pageWidth =
                VisioShapeSheet.GetNumber(page.PageSheet, "PageWidth");
            double pageHeight =
                VisioShapeSheet.GetNumber(page.PageSheet, "PageHeight");

            VisioShapeSheet.SetNumber(shape, "PinX",
                bounds[0].getRelative2DPosX() * pageWidth);
            VisioShapeSheet.SetNumber(shape, "PinY",
                bounds[0].getRelative2DPosY() * pageHeight);
            VisioShapeSheet.SetNumber(shape, "Width",
                bounds[1].getRelative2DPosX() * pageWidth);
            VisioShapeSheet.SetNumber(shape, "Height",
                bounds[1].getRelative2DPosY() * pageHeight);
        }
    }
}
