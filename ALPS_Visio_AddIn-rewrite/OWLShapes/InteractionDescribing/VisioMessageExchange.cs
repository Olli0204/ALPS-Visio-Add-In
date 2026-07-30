using alps.net.api.ALPS;
using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioMessageExchange : MessageExchange, IVisioExportableWithShape
    {
        private const string shapeType = Constants.SIDMasters.StandardMessageConnector;

        private readonly IShapeExport export;
        public VisioMessageExchange(IModelLayer layer) : base(layer) { export = new PASSProcessModelElementExport(this); }
        protected VisioMessageExchange() { export = new PASSProcessModelElementExport(this); }
        public void ExportToVisio(Visio.Page page)
        {
            // A message exchange is normally exported by its MessageExchangeList.
            // Do not create a second connector when the model layer reaches the
            // same exchange again as an individual element.
            if (GetShape() != null) return;

            export.Export(shapeType, page, VH.GetBounds(this));

            // Keep the semantic endpoints on the connector itself. Some
            // message connector masters temporarily lose one glue entry while
            // Visio lays out a coordinate-free import. The common auto-arrange
            // pipeline can restore both ends deterministically from these IDs.
            ISubject sender = getSender();
            ISubject receiver = getReceiver();
            if (sender != null)
            {
                VH.SetProperty(GetShape(),
                    Constants.Properties.MessageExchange.OriginSubject,
                    sender.getModelComponentID());
            }
            if (receiver != null)
            {
                VH.SetProperty(GetShape(),
                    Constants.Properties.MessageExchange.TargetSubject,
                    receiver.getModelComponentID());
            }

            // set path (auto arrange)
            if (sender is IVisioExportableWithShape exportableSender && exportableSender.GetShape() != null)
                this.GetShape().CellsU["BeginX"].GlueToPos(exportableSender.GetShape(), 1, 0.5);
            if (receiver is IVisioExportableWithShape exportableReceiver && exportableReceiver.GetShape() != null)
                this.GetShape().CellsU["EndX"].GlueToPos(exportableReceiver.GetShape(), 0, 0.5);

            // TODO: AbstractMessageExchange
            // TODO: FinalizedMessageExchange -> alps.net.api
        }

        public bool PrepareDimensions() // TODO: prepare dimensions
        {
            return false; // routing follow points
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioMessageExchange();
        }

        public Visio.Shape GetShape()
        {
            return export.GetShape();
        }
    }
}
