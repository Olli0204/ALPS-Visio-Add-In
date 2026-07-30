namespace PassBpmnConverter.Bpmn;

public interface IFormalExpression : IExpression
{
}

[BpmnType("formalExpression", BpmnModelConstants.BpmnNs)]
public class FormalExpression : Expression, IFormalExpression
{
}
