namespace PassBpmnConverter.Bpmn;

public interface IMessageFlow : IBaseElement
{
    string? Name { get; set; }

    IInteractionNode SourceRef { get; set; }

    IInteractionNode TargetRef { get; set; }

    IMessage? MessageRef { get; set; }
}

[BpmnType("messageFlow", BpmnModelConstants.BpmnNs)]
public class MessageFlow : BaseElement, IMessageFlow
{
    [BpmnAttribute("name")]
    public string? Name { get; set; }

    [BpmnAttribute("sourceRef")]
    public IInteractionNode SourceRef { get; set; }

    [BpmnAttribute("targetRef")]
    public IInteractionNode TargetRef { get; set; }

    [BpmnAttribute("messageRef")]
    public IMessage? MessageRef { get; set; }
}
