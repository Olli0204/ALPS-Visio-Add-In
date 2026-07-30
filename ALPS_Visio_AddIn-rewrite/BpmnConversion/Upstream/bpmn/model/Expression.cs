namespace PassBpmnConverter.Bpmn;

// TODO: according to https://www.omg.org/spec/BPMN/20100501/Semantic.xsd this inherit tBaseElementWithMixedContent
public interface IExpression : IBaseElement
{
    string? Body { get; set; }
}

[BpmnType("expression", BpmnModelConstants.BpmnNs)]
public class Expression : BaseElement, IExpression
{
    public string? Body { get; set; }
}
