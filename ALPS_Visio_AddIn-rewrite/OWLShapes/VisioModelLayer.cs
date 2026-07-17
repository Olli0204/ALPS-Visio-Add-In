using alps.net.api.ALPS;
using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using alps.net.api.util;
using System.Collections.Generic;
using System.Linq;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioModelLayer : ModelLayer, IVisioExportable
    {
        public VisioModelLayer(IPASSProcessModel model, string labelForID = null, string comment = null, string additionalLabel = null, IList<IIncompleteTriple> additionalAttribute = null) : base(model, labelForID, comment, additionalLabel, additionalAttribute) { }
        protected VisioModelLayer() { }

        public void ExportToVisio(Visio.Page page)
        {
            SetPageDimensions(page);

            // hasPriorityNumber
            VH.SetProperty(page.PageSheet, Constants.Properties.PriorityOrderNumber, this.priorityNumber.ToString());

            int elementIndex = 0;
            foreach (IPASSProcessModelElement modelElement in this.getElements().Values
                .Where(element => !(element is ISubjectBehavior))
                .OrderBy(GetExportOrder))
            {
                if (!(modelElement is IVisioExportable exportable)) continue;

                if (exportable is IVisioExportableWithShape shapeExportable)
                    VisioLayout.PrepareOrArrange(shapeExportable, elementIndex++, false);

                exportable.ExportToVisio(page);
            }
        }

        private static int GetExportOrder(IPASSProcessModelElement element)
        {
            if (element is ISubject) return 0;
            if (element is IMessageExchangeList) return 1;
            if (element is IMessageExchange) return 2;
            return 3;
        }

        /// <summary>
        /// Calculate dimensions for this model and apply to given page.
        /// 
        /// Note: This is inconsistent, it would be great to add some size to the standard.
        /// </summary>
        private void SetPageDimensions(Visio.Page page)
        {
            double pageRatio = 1;
            double sumWidth = 0;
            int subjectCount = 0;
            foreach (ISubject modelElement in this.getElements().Select(x => x.Value).OfType<ISubject>())
            {
                if (modelElement is ISystemInterfaceSubject) continue;

                if ((modelElement is IFullySpecifiedSubject || modelElement is IInterfaceSubject))
                {
                    pageRatio = modelElement.get2DPageRatio();
                    double width = modelElement.getRelative2DWidth();
                    sumWidth += width;
                    if (width > 0) subjectCount++;
                }
            }
            if (subjectCount == 0 || sumWidth <= 0 || pageRatio <= 0)
            {
                VH.SetSizeMM(page.PageSheet, "PageWidth", 297);
                VH.SetSizeMM(page.PageSheet, "PageHeight", 210);
                return;
            }

            double averageWidth = sumWidth / subjectCount;

            // the average subject is 32 mm wide
            double newPageWidth = 32 / averageWidth + 1;
            double newPageHeight = newPageWidth / pageRatio;

            // FEAT: round to nearest A4 page

            VH.SetSizeMM(page.PageSheet, "PageWidth", newPageWidth);
            VH.SetSizeMM(page.PageSheet, "PageHeight", newPageHeight);
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioModelLayer();
        }
    }
}
