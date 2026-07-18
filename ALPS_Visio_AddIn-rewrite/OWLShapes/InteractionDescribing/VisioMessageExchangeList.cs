using alps.net.api.ALPS;
using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;
using Visio = Microsoft.Office.Interop.Visio;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioMessageExchangeList : MessageExchangeList, IVisioExportable
    {
        public VisioMessageExchangeList(IModelLayer layer) : base(layer) { }
        protected VisioMessageExchangeList() { }

        public void ExportToVisio(Visio.Page page)
        {
            if (this.getMessageExchanges().Values.FirstOrDefault() is IVisioExportableWithShape messageExchangeWithConnector)
            {
                VisioLayout.PrepareOrArrange(messageExchangeWithConnector, 0, false);

                // store previous shapes
                List<Visio.Shape> previousShapes = new List<Visio.Shape>();
                foreach (Visio.Shape shape in page.Shapes) previousShapes.Add(shape);

                messageExchangeWithConnector.ExportToVisio(page);

                // find message box
                Visio.Shape messageBox = null;
                foreach (Visio.Shape shape in page.Shapes)
                    if (shape.CellExistsU["User.idOnPage", 0] != 0 &&
                        shape.CellsU["User.idOnPage"].Result[""] == messageExchangeWithConnector.GetShape().CellsU["User.idOfCorrespondingShape"].Result[""])
                    {
                        messageBox = shape;
                        break;
                    }

                if (messageBox == null)
                    messageBox = CreateMessageContainer(page, messageExchangeWithConnector.GetShape());

                // delete wrong shapes
                // alternative idea: delete messages with default label (or label == id)
                List<Visio.Shape> shapesToDelete = new List<Visio.Shape>();
                foreach (Visio.Shape shape in page.Shapes)
                    if (!previousShapes.Contains(shape) && shape != messageExchangeWithConnector.GetShape() && shape != messageBox)
                        shapesToDelete.Add(shape);

                foreach (Visio.Shape shape in shapesToDelete)
                    shape.Delete();

                // aggregate list
                foreach (IMessageExchange messageExchange in this.getMessageExchanges().Values)
                {
                    if (messageExchange.getMessageType() is IVisioExportableWithShape exportable)
                    {
                        VisioLayout.PrepareOrArrange(exportable, 0, false);
                        exportable.ExportToVisio(page);

                        messageBox.ContainerProperties.InsertListMember(exportable.GetShape(), 0);
                        exportable.GetShape().BringToFront();
                    }
                }

                CenterMessageContainer(messageBox, messageExchangeWithConnector.GetShape());
            }
        }

        private static Visio.Shape CreateMessageContainer(Visio.Page page, Visio.Shape connector)
        {
            Visio.Shape messageBox = VH.Place(Constants.SIDMasters.MessageBox, page);

            string connectorId = connector.ID.ToString(CultureInfo.InvariantCulture);
            string messageBoxId = messageBox.ID.ToString(CultureInfo.InvariantCulture);
            string connectorReference = "Sheet." + connectorId;
            string messageBoxReference = "Sheet." + messageBoxId;

            connector.CellsU["User.idOfCorrespondingShape"].FormulaU =
                "=" + messageBoxReference + "!User.idOnPage";
            connector.CellsU["User.globalX"].FormulaU =
                "=SETATREF(" + messageBoxReference + "!User.connectorControlsPositionX)";
            connector.CellsU["User.globalY"].FormulaU =
                "=SETATREF(" + messageBoxReference + "!User.connectorControlsPositionY)";
            messageBox.CellsU["User.idOfCorrespondingShape"].FormulaU =
                "=" + connectorReference + "!User.idOnPage";

            return messageBox;
        }

        private static void CenterMessageContainer(Visio.Shape messageBox, Visio.Shape connector)
        {
            double centerX = (VH.GetSize(connector, "BeginX") + VH.GetSize(connector, "EndX")) / 2d;
            double centerY = (VH.GetSize(connector, "BeginY") + VH.GetSize(connector, "EndY")) / 2d;

            VH.SetSize(messageBox, "PinX", centerX);
            VH.SetSize(messageBox, "PinY", centerY);
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioMessageExchangeList();
        }
    }
}
