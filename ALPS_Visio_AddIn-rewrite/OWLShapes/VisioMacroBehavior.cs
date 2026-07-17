using System.Collections.Generic;
using System.Linq;
using alps.net.api.ALPS;
using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using alps.net.api.util;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioMacroBehavor : MacroBehavior, IVisioExportable
    {
        public VisioMacroBehavor(IModelLayer layer, string labelForID = null, ISubject subject = null, ISet<IBehaviorDescribingComponent> behaviorDescribingComponents = null, ISet<IStateReference> stateReferences = null, IState initialStateOfBehavior = null, int priorityNumber = 0, string comment = null, string additionalLabel = null, IList<IIncompleteTriple> additionalAttribute = null) : base(layer, labelForID, subject, behaviorDescribingComponents, stateReferences, initialStateOfBehavior, priorityNumber, comment, additionalLabel, additionalAttribute) { }

        protected VisioMacroBehavor() { }

        public void ExportToVisio(Visio.Page currentPage)
        {
            bool usesFallbackLayout = VisioLayout.ArrangeBehavior(this.getBehaviorDescribingComponents().Values);
            if (usesFallbackLayout)
            {
                VH.SetSizeMM(currentPage.PageSheet, "PageWidth", 420);
                VH.SetSizeMM(currentPage.PageSheet, "PageHeight", 240);
                VH.ConfigureFallbackSbdRouting(currentPage);
            }

            int componentIndex = 0;
            foreach (IBehaviorDescribingComponent component in this.getBehaviorDescribingComponents().Values
                .OrderBy(component => component is ITransition))
            {
                if (!(component is IVisioExportable exportable)) continue;

                if (exportable is IVisioExportableWithShape shapeExportable)
                    VisioLayout.PrepareOrArrange(shapeExportable, componentIndex++, true);

                if (exportable is IState || exportable is ITransition) exportable.ExportToVisio(currentPage);
            }

        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioMacroBehavor();
        }
    }
}
