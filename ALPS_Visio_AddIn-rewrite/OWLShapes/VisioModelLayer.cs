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
        public VisioModelLayer(IPASSProcessModel model, string labelForID = null, string comment = null, string additionalLabel = null, IList<IPASSTriple> additionalAttribute = null) : base(model, labelForID, comment, additionalLabel, additionalAttribute) { }
        protected VisioModelLayer() { }

        public void ExportToVisio(Visio.Page page)
        {
            bool usesFallbackLayout = VisioLayout.ArrangeModelLayer(this.getElements().Values);
            ModelPageSizer.Apply(page,
                this.getElements().Values.OfType<ISubject>(), usesFallbackLayout);
            ISet<string> exchangesExportedByList = new HashSet<string>(
                this.getElements().Values
                    .OfType<IMessageExchangeList>()
                    .SelectMany(list => list.getMessageExchanges().Values)
                    .Where(exchange => exchange != null)
                    .Select(exchange => exchange.getModelComponentID()));

            // hasPriorityNumber
            VH.SetProperty(page.PageSheet, Constants.Properties.PriorityOrderNumber, this.priorityNumber.ToString());

            int elementIndex = 0;
            foreach (IPASSProcessModelElement modelElement in this.getElements().Values
                .Where(element => !(element is ISubjectBehavior))
                .Where(element => !IsContainedInMessageExchangeList(element, exchangesExportedByList))
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

        private static bool IsContainedInMessageExchangeList(IPASSProcessModelElement element, ISet<string> exchangesExportedByList)
        {
            return element is IMessageExchange exchange
                && exchangesExportedByList.Contains(exchange.getModelComponentID());
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioModelLayer();
        }
    }
}
