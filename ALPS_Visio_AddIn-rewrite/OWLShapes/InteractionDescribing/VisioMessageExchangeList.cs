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

                // center message box
                messageBox.CellsU["Actions.Center.Action"].Trigger();

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
            }
        }

        private static Visio.Shape CreateMessageContainer(Visio.Page page, Visio.Shape connector)
        {
            Visio.Shape messageBox = VH.Place(Constants.SIDMasters.MessageBox, page);

            string connectorId = connector.ID.ToString(CultureInfo.InvariantCulture);
            string messageBoxId = messageBox.ID.ToString(CultureInfo.InvariantCulture);

            connector.CellsU["User.idOfCorrespondingShape"].FormulaU = messageBoxId;
            messageBox.CellsU["User.idOfCorrespondingShape"].FormulaU = connectorId;

            return messageBox;
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioMessageExchangeList();
        }
    }
}
