using System.Collections.Generic;
using alps.net.api.ALPS;
using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using alps.net.api.util;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioSubjectBehavior : SubjectBehavior, IVisioExportable
    {
        public VisioSubjectBehavior(IModelLayer layer, string labelForID = null, ISubject subject = null, ISet<IBehaviorDescribingComponent> behaviorDescribingComponents = null, IState initialStateOfBehavior = null, int priorityNumber = 0, string comment = null, string additionalLabel = null, IList<IPASSTriple> additionalAttribute = null) : base(layer, labelForID, subject, behaviorDescribingComponents, initialStateOfBehavior, priorityNumber, comment, additionalLabel, additionalAttribute) { }
        protected VisioSubjectBehavior() { }

        public void ExportToVisio(Visio.Page currentPage)
        {
            BehaviorExport.Export(getBehaviorDescribingComponents().Values, currentPage);
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioSubjectBehavior();
        }
    }
}
