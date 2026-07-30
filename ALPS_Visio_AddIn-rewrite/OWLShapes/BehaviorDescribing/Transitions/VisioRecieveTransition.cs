using alps.net.api.parsing;
using alps.net.api.StandardPASS;
using alps.net.api.util;
using Visio = Microsoft.Office.Interop.Visio;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;
using System.Collections.Generic;
using System.Linq;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class VisioReceiveTransition : ReceiveTransition, IVisioExportableWithShape
    {
        private const string shapeType = Constants.SBDMasters.ReceiveTransition;
        
        private readonly IShapeExport export;
        public VisioReceiveTransition(IState sourceState, IState targetState, string labelForID = null, ITransitionCondition transitionCondition = null, ITransition.TransitionType transitionType = ITransition.TransitionType.Standard, ISet<IDataMappingIncomingToLocal> dataMappingIncomingToLocal = null, int priorityNumber = 0, string comment = null, string additionalLabel = null, IList<IPASSTriple> additionalAttribute = null) : base(sourceState, targetState, labelForID, transitionCondition, transitionType, dataMappingIncomingToLocal, priorityNumber, comment, additionalLabel, additionalAttribute) { export = new TransitionExport(this); }
        protected VisioReceiveTransition() { export = new TransitionExport(this); }

        public void ExportToVisio(Visio.Page page)
        {
            export.Export(shapeType, page, VH.GetBounds(this));

            var condition = getTransitionCondition();
            if (condition == null) return;

            // sender
            ISubject sender = condition.getMessageSentFrom();
            if (sender != null && sender.getModelComponentLabels().Count > 0)
            {
                VH.SetUser(export.GetShape(), Constants.Properties.Transition.ReceiverSenderListForSubject, ";" + sender.getModelComponentLabelsAsStrings()[0]);
                VH.SetUser(export.GetShape(), Constants.Properties.Transition.ReceiverSenderListForSubjectID, ";" + sender.getModelComponentID());
                VH.SetPropertyFormulaU(export.GetShape(), Constants.Properties.Transition.MessageSender, "INDEX(1, Prop.senderOfMessage.Format)");
            }

            // message
            IMessageSpecification messageSpec = condition.getReceptionOfMessage();
            if (messageSpec != null && messageSpec.getModelComponentLabels().Count > 0)
            {
                VH.SetUser(export.GetShape(), Constants.Properties.Transition.PossibleMessageList, ";" + messageSpec.getModelComponentLabelsAsStrings()[0]);
                VH.SetUser(export.GetShape(), Constants.Properties.Transition.PossibleMessageListID, ";" + messageSpec.getModelComponentID());
                VH.SetPropertyFormulaU(export.GetShape(), Constants.Properties.Transition.Message, "INDEX(1, Prop.Message.Format)");
            }

            VH.SetProperty(export.GetShape(), Constants.Properties.Transition.MultiReceiveLowerBound, condition.getMultipleLowerBound().ToString());
            VH.SetProperty(export.GetShape(), Constants.Properties.Transition.MultiReceiveUpperBound, condition.getMultipleUpperBound().ToString());

            // priority number
            VH.SetProperty(export.GetShape(), Constants.Properties.Transition.AlternativePriorityNumber, getPriorityNumber().ToString());

            VH.SetPropertyFormulaU(export.GetShape(), Constants.Properties.Transition.ReceiveType,
                "INDEX(" + (int)condition.getReceiveType() + ", Prop.receiveType.Format)");

            // add data mapping
            if (getDataMappingFunctions().Count > 0)
            {
                List<IDataMappingIncomingToLocal> tempList = getDataMappingFunctions().Values.ToList();
                if (tempList.Count > 0)
                {
                    string dataMappingString = tempList[0].getDataMappingString();
                    VH.SetProperty(export.GetShape(), Constants.Properties.Transition.DataMappingIncomming, dataMappingString);
                }
            }
        }

        public bool PrepareDimensions() // TODO: prepare dimensions
        {
            return false;
        }

        public override IParseablePASSProcessModelElement getParsedInstance()
        {
            return new VisioReceiveTransition();
        }

        public Visio.Shape GetShape()
        {
            return export.GetShape();
        }
    }
}
