using alps.net.api.ALPS;
using alps.net.api.parsing;
using Visio = Microsoft.Office.Interop.Visio;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioCommunicationChannel : CommunicationChannel, IVisioExportableWithShape
    {
        private readonly IShapeExport export;

        public VisioCommunicationChannel(IModelLayer layer) : base(layer)
        {
            export = new PASSProcessModelElementExport(this);
        }

        protected VisioCommunicationChannel()
        {
            export = new PASSProcessModelElementExport(this);
        }

        public void ExportToVisio(Visio.Page page)
        {
            export.Export(Constants.SIDMasters.AbstractCommunicationChannel, page, VH.GetBounds(this));
        }

        public bool PrepareDimensions()
        {
            return false;
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioCommunicationChannel();
        }

        public Visio.Shape GetShape()
        {
            return export.GetShape();
        }
    }
}
