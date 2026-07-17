using alps.net.api.ALPS;
using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using Visio = Microsoft.Office.Interop.Visio;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioSystemInterfaceSubject : SystemInterfaceSubject, IVisioExportableWithShape
    {
        private readonly IShapeExport export;

        public VisioSystemInterfaceSubject(IModelLayer layer) : base(layer)
        {
            export = new SubjectExport(this);
        }

        protected VisioSystemInterfaceSubject()
        {
            export = new SubjectExport(this);
        }

        public void ExportToVisio(Visio.Page page)
        {
            export.Export(Constants.SIDMasters.SystemInterfaceSubject, page, VH.GetBounds(this));
        }

        public bool PrepareDimensions()
        {
            return false;
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioSystemInterfaceSubject();
        }

        public Visio.Shape GetShape()
        {
            return export.GetShape();
        }
    }
}
