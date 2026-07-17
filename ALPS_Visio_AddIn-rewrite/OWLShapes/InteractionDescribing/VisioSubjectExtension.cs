using alps.net.api.ALPS;
using alps.net.api.parsing;
using Visio = Microsoft.Office.Interop.Visio;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioSubjectExtension : SubjectExtension, IVisioExportableWithShape
    {
        private readonly IShapeExport export;

        public VisioSubjectExtension(IModelLayer layer) : base(layer)
        {
            export = new SubjectExport(this);
        }

        protected VisioSubjectExtension()
        {
            export = new SubjectExport(this);
        }

        public void ExportToVisio(Visio.Page page)
        {
            export.Export(Constants.SIDMasters.ActorExtension, page, VH.GetBounds(this));
        }

        public bool PrepareDimensions()
        {
            return false;
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioSubjectExtension();
        }

        public Visio.Shape GetShape()
        {
            return export.GetShape();
        }
    }
}
