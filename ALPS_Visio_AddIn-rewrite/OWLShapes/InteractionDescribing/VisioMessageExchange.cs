using alps.net.api.ALPS;
using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using ALPS_Visio_AddIn_rewrite.VisioInfrastructure;
using System.Runtime.InteropServices;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;
using Visio = Microsoft.Office.Interop.Visio;
using LayoutDirection =
    ALPS_Visio_AddIn_rewrite.VisioHelper.GraphLayoutDirection;

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
            RebindToSemanticEndpoints();

            // TODO: AbstractMessageExchange
            // TODO: FinalizedMessageExchange -> alps.net.api
        }

        internal bool RebindToSemanticEndpoints()
        {
            Visio.Shape connector = GetShape();
            if (connector == null) return false;

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

            // Store the concrete page shape IDs as well as the semantic IDs.
            // The SID master can remove a Glue entry when its coupled Message
            // Box moves; a page-local ID restores the exact original shape
            // without depending on the master's Shape Data formulas.
            bool sourceBound = false;
            if (sender is IVisioExportableWithShape exportableSender
                && exportableSender.GetShape() != null)
            {
                Visio.Shape senderShape = exportableSender.GetShape();
                VH.SetUser(GetShape(),
                    Constants.UserCells.AutoArrangeSourceShapeId,
                    senderShape.ID);
                sourceBound = TryGlueEndpoint(
                    connector, senderShape, true, 1, 0.5);
            }

            bool targetBound = false;
            if (receiver is IVisioExportableWithShape exportableReceiver
                && exportableReceiver.GetShape() != null)
            {
                Visio.Shape receiverShape = exportableReceiver.GetShape();
                VH.SetUser(GetShape(),
                    Constants.UserCells.AutoArrangeTargetShapeId,
                    receiverShape.ID);
                targetBound = TryGlueEndpoint(
                    connector, receiverShape, false, 0, 0.5);
            }

            return sourceBound && targetBound;
        }

        internal bool RebindToSemanticEndpoints(
            Visio.IVPage page, LayoutDirection direction)
        {
            Visio.Shape connector = GetShape();
            // Auto-Arrange renders SID message exchanges through a native
            // visual connector. The original grouped stencil master remains
            // hidden solely as the semantic owner of Shape Data and the
            // MessageBox relationship; rebinding it would revive the broken
            // internal leader geometry that the visual connector replaces.
            if (VisioSidMessageConnectorRenderer.IsSemanticShadow(connector))
                return true;

            IVisioExportableWithShape exportableSender =
                getSender() as IVisioExportableWithShape;
            IVisioExportableWithShape exportableReceiver =
                getReceiver() as IVisioExportableWithShape;
            if (connector == null
                || exportableSender == null
                || exportableReceiver == null
                || exportableSender.GetShape() == null
                || exportableReceiver.GetShape() == null)
            {
                return false;
            }

            Visio.Shape senderShape = exportableSender.GetShape();
            Visio.Shape receiverShape = exportableReceiver.GetShape();
            if (VisioConnectorRebinder.AreSemanticEndpointsBound(
                connector, senderShape, receiverShape))
            {
                return true;
            }

            bool rebound = VisioConnectorRebinder.RebindKnownConnector(
                page, connector, senderShape, receiverShape, direction);
            return rebound
                && VisioConnectorRebinder.AreSemanticEndpointsBound(
                    connector, senderShape, receiverShape);
        }

        private static bool TryGlueEndpoint(Visio.Shape connector,
            Visio.Shape endpointShape, bool isSource,
            double relativeX, double relativeY)
        {
            // Preserve the BeginX/EndY endpoint convention of the grouped SID
            // stencil master; its legacy geometry was authored with these cells.
            string endpointCellName = isSource ? "BeginX" : "EndY";
            try
            {
                connector.CellsU[endpointCellName].GlueToPos(
                    endpointShape, relativeX, relativeY);
                if (IsEndpointGluedTo(
                    connector, endpointShape, isSource))
                {
                    return true;
                }
            }
            catch (COMException)
            {
                // Retry below with Visio's dynamic glue target.
            }

            try
            {
                connector.CellsU[endpointCellName]
                    .GlueTo(endpointShape.CellsU["PinX"]);
                return IsEndpointGluedTo(
                    connector, endpointShape, isSource);
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static bool IsEndpointGluedTo(Visio.Shape connector,
            Visio.Shape endpointShape, bool isSource)
        {
            int minimumFromPart = isSource ? 7 : 10;
            int maximumFromPart = isSource ? 9 : 12;
            try
            {
                foreach (Visio.Connect connection in connector.Connects)
                {
                    if (connection.FromPart >= minimumFromPart
                        && connection.FromPart <= maximumFromPart
                        && connection.ToSheet != null
                        && connection.ToSheet.ID == endpointShape.ID)
                    {
                        return true;
                    }
                }
            }
            catch (COMException)
            {
                return false;
            }

            return false;
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
