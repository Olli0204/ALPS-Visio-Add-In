using System.Collections.Generic;
using alps.net.api.ALPS;
using alps.net.api.StandardPASS;
using Visio = Microsoft.Office.Interop.Visio;
using VH = ALPS_Visio_AddIn_rewrite.VisioHelper;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes
{
    public class TransitionExport : PASSProcessModelElementExport
    {
		private readonly ITransition transition;

        /// <summary>
        /// Shape export for transition
        /// </summary>
        public TransitionExport(ITransition transition) : base(transition)
        {
            this.transition = transition;
        }

        public override void Export(string shapeType, Visio.Page page, IList<ISimple2DVisualizationPoint> bounds)
        {
            base.Export(shapeType, page, bounds);

            // TODO: hasSourceState only State @ originState
            // TODO: hasTargetState only State @ targetState
            // TODO: hasTransitionCondition exactly 1 TransitionCondition
            // -> hasToolSpecificDefinition max 1 string

            // DoTransition
            // TODO: hasPriorityNumber max 1 int(>=0)
            // -> DoTransitionCondition
            // label mit condition string überschreiben

            // CommunicationTransition (contains Receive and Send)
            // -> MessageExchangeCondition (requiresPerformedMessageExchange exactly 1 MessageExchange)

            // ReceiveTransition
            // TODO: hasDataMappingFunction min 1 DataMappingIncomingToLocal (hasFeelExpressionAsDataMapping OR hasToolSpecificDefinition exactly 1 string, hasDataMappingString exactly 1 string)
            // TODO: hasPriorityNumber max 1 int(>=0)
            // -> ReceiveTransitionCondition (hasMultiReceiveLower/UpperBound max 1 int(>0), hasReceiveType max 1 ReceiveType(MultiReceiveFromAllKnown, MultiReceiveFromKnown, Standard), requiresMessageSentFrom max 1 Subject, requiresReceptionOfMessage max 1 MessageSpecification)

            // SendTransition
            // TODO: hasDataMappingFunction min 1 DataMappingLocalToOutgoing (hasFeelExpressionAsDataMapping OR hasToolSpecificDefinition exactly 1 string, hasDataMappingString exactly 1 string)
            // -> SendTransitionCondition (hasMultiSendLower/UpperBound max 1 int(>0), hasSendType max 1 SendType(MultiSendToAll, MultiSendToKnown, MultiSendToNew, Standard), requiresMessageSentTo max 1 Subject, requiresSendingOfMessage max 1 MessageSpecification)

            // SendingFailedTransition
            // TODO: hasSourceState only SendState @ originState
            // -> SendingFailedCondition

            // UserCancelTransition

            // TimeTransition (Reminder (CalendarBased, TimeBased) hasSourceState max 0 SendState, Timer (BusinessDay, DayTime, YearMonth))
            // -> TimeTransitionCondition (hasTimeValue exactly 1 string or (DayTimeTimer)dateTime or (YearMonthTimer)yearMonthDuration)

            // TODO: AbstractPASSTransition (AdviceTransitionType, FinalizedTransitionType, PrecedenceTransitionType, TriggerTransitionType, FlowRestrictor)
            // modelComponentType

            //////////////////////////////////////////////////////////
            // TODO: in visio:
            // Sender of Message: senderOfMessage
            // Message to be Received/Sent: message
            //((ReceiveTransition)transition).getTransitionCondition().getRequiresPerformedMessageExchange()
            // Receive Type: receiveType
            // Lower Bound if Multi Receive: multiReceiveLowerBound
            // Upper Bound if Multi Receive: multiReceiveUpperBound
            // Priority of Message Receive: alternativePriorityNumber
            // Type of Time based Transition: timeoutType
            // timeOutTime: timeOutTime
            // time Date/Frequency: timeOutDate
            // implements/represents: implements
            // Receiver of Message: receivingSubject
            // Sending Type: sendingType
            // Lower Bound if Multi Send: multiSendLowerBound
            // Upper Bound if Multi Send: multiSendUpperBound

            // set path (auto arrange)
            VisioLayout.GetTransitionPorts(transition, out double sourcePortY, out double targetPortY);
            if (transition.getSourceState() is IVisioExportableWithShape exportableSender && exportableSender.GetShape() != null)
                this.GetShape().CellsU["BeginX"].GlueToPos(exportableSender.GetShape(), 1, sourcePortY);
            if (transition.getTargetState() is IVisioExportableWithShape exportableReceiver && exportableReceiver.GetShape() != null)
                this.GetShape().CellsU["EndX"].GlueToPos(exportableReceiver.GetShape(), 0, targetPortY);
            PositionLabel(sourcePortY, targetPortY);

            // set box movement
            VH.SetProperty(shape, Constants.Properties.Transition.BoxCanBeMovedFreely, "FALSE");

            // set implements
            if (transition.getImplementedInterfaces().Count > 0)
                VH.SetProperty(shape, Constants.Properties.Transition.Implements,
                    string.Join(";", transition.getImplementedInterfaces().Keys));
        }

        private void PositionLabel(double sourcePortY, double targetPortY)
        {
            Visio.Shape transitionShape = this.GetShape();
            if (transitionShape == null
                || transitionShape.CellExistsU["TxtPinX", 0] == 0
                || transitionShape.CellExistsU["TxtPinY", 0] == 0) return;

            // Place the label close to the source and on the free side of the
            // connector. Branches therefore no longer stack their labels at the
            // centre of the diagram.
            transitionShape.CellsU["TxtPinX"].FormulaU = "Width*0.30";
            double averagePortY = (sourcePortY + targetPortY) / 2.0;
            transitionShape.CellsU["TxtPinY"].FormulaU = averagePortY <= 0.5
                ? "Height*0.5 + 0.18 in"
                : "Height*0.5 - 0.18 in";
        }
    }
}
