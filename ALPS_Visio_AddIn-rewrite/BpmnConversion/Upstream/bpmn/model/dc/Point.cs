namespace PassBpmnConverter.Bpmn.DC;

public interface IPoint
{
    double X { get; set; }
    double Y { get; set; }
}

[BpmnType("Point", BpmnModelConstants.OmgDcNs)]
public class Point : IPoint
{
    [BpmnAttribute("x")]
    public double X { get; set; }

    [BpmnAttribute("y")]
    public double Y { get; set; }
}
