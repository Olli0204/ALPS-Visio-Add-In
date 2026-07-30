using alps.net.api.ALPS;
using alps.net.api.StandardPASS;
using System;
using System.Collections.Generic;
using System.Linq;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    /// <summary>
    /// Exports the shared state/transition structure of subject and macro behaviors.
    /// </summary>
    internal static class BehaviorExport
    {
        public static void Export(
            IEnumerable<IBehaviorDescribingComponent> behaviorComponents,
            Visio.Page page)
        {
            if (behaviorComponents == null)
                throw new ArgumentNullException(nameof(behaviorComponents));
            if (page == null) throw new ArgumentNullException(nameof(page));

            List<IBehaviorDescribingComponent> components =
                behaviorComponents.ToList();
            bool usesFallbackLayout = VisioLayout.ArrangeBehavior(components);
            if (usesFallbackLayout)
            {
                VH.SetSizeMM(page.PageSheet, "PageWidth", 420);
                VH.SetSizeMM(page.PageSheet, "PageHeight", 210);
                VH.ConfigureFallbackSbdRouting(page);
            }

            int componentIndex = 0;
            foreach (IBehaviorDescribingComponent component in components
                .OrderBy(candidate => candidate is ITransition))
            {
                IVisioExportable exportable = component as IVisioExportable;
                if (exportable == null) continue;

                IVisioExportableWithShape shapeExportable =
                    exportable as IVisioExportableWithShape;
                if (shapeExportable != null)
                    VisioLayout.PrepareOrArrange(
                        shapeExportable, componentIndex++, true);

                if (exportable is IState || exportable is ITransition)
                    exportable.ExportToVisio(page);
            }

            if (usesFallbackLayout)
            {
                VH.AutoArrangePage(
                    page, VH.GraphLayoutDirection.TopDown);
            }
        }
    }
}
