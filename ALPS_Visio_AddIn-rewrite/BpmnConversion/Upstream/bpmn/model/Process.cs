
namespace PassBpmnConverter.Bpmn;

// TODO: add additional attributes/elements (e.g., isClosed, isExecutable)
public interface IProcess : ICallableElement, IFlowElementsContainer
{
    bool IsExecutable { get; set; }
}

[BpmnType("process", BpmnModelConstants.BpmnNs)]
public class Process : CallableElement, IProcess
{
    [BpmnAttribute("isExecutable")]
    public bool IsExecutable { get; set; } = false;

    [BpmnElement]
    public IList<IFlowElement> FlowElements { get; set; } = new List<IFlowElement>();
}
