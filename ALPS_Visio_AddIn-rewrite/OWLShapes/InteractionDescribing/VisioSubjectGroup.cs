using alps.net.api.ALPS;
using alps.net.api.parsing;
using Visio = Microsoft.Office.Interop.Visio;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioSubjectGroup : SubjectGroup, IVisioExportableWithShape
    {
        private readonly IShapeExport export;

        public VisioSubjectGroup(IModelLayer layer) : base(layer)
        {
            export = new SubjectExport(this);
        }

        protected VisioSubjectGroup()
        {
            export = new SubjectExport(this);
        }

        public void ExportToVisio(Visio.Page page)
        {
            export.Export(Constants.SIDMasters.SubjectGroup, page, VH.GetBounds(this));
        }

        public bool PrepareDimensions()
        {
            return false;
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioSubjectGroup();
        }

        public Visio.Shape GetShape()
        {
            return export.GetShape();
        }
    }
}
