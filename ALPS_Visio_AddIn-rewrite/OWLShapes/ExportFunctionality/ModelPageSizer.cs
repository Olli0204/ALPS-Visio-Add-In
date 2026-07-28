using alps.net.api.ALPS;
using alps.net.api.StandardPASS;
using System;
using System.Collections.Generic;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    /// <summary>
    /// Derives a readable SID page size from available visualization metadata.
    /// </summary>
    internal static class ModelPageSizer
    {
        private const double DefaultWidthMillimeters = 297;
        private const double DefaultHeightMillimeters = 210;
        private const double AverageSubjectWidthMillimeters = 32;

        public static void Apply(Visio.Page page, IEnumerable<ISubject> subjects,
            bool usesFallbackLayout)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (subjects == null) throw new ArgumentNullException(nameof(subjects));

            if (usesFallbackLayout)
            {
                SetDefaultSize(page);
                return;
            }

            double pageRatio = 1;
            double sumWidth = 0;
            int subjectCount = 0;
            foreach (ISubject subject in subjects)
            {
                if (subject is ISystemInterfaceSubject) continue;
                if (!(subject is IFullySpecifiedSubject)
                    && !(subject is IInterfaceSubject))
                    continue;

                pageRatio = subject.get2DPageRatio();
                double width = subject.getRelative2DWidth();
                sumWidth += width;
                if (width > 0) subjectCount++;
            }

            if (subjectCount == 0 || sumWidth <= 0 || pageRatio <= 0)
            {
                SetDefaultSize(page);
                return;
            }

            double averageWidth = sumWidth / subjectCount;
            double pageWidth =
                AverageSubjectWidthMillimeters / averageWidth + 1;
            double pageHeight = pageWidth / pageRatio;
            VH.SetSizeMM(page.PageSheet, "PageWidth", pageWidth);
            VH.SetSizeMM(page.PageSheet, "PageHeight", pageHeight);
        }

        private static void SetDefaultSize(Visio.Page page)
        {
            VH.SetSizeMM(page.PageSheet, "PageWidth", DefaultWidthMillimeters);
            VH.SetSizeMM(page.PageSheet, "PageHeight", DefaultHeightMillimeters);
        }
    }
}
