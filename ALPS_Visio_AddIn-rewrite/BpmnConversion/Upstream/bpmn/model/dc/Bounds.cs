namespace PassBpmnConverter.Bpmn.DC;

public interface IBounds
{
    double X { get; set; }

    double Y { get; set; }

    double Width { get; set; }

    double Height { get; set; }
}

[BpmnType("Bounds", BpmnModelConstants.OmgDcNs)]
public class Bounds : IBounds
{
    [BpmnAttribute("x")]
    public double X { get; set; }

    [BpmnAttribute("y")]
    public double Y { get; set; }

    [BpmnAttribute("width")]
    public double Width { get; set; }

    [BpmnAttribute("height")]
    public double Height { get; set; }
}
