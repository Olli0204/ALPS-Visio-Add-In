namespace ALPS_Visio_AddIn_rewrite.NlpChecking
{
    internal sealed class NlpTrainingExample
    {
        public string Name { get; set; }
        public string ShapeType { get; set; }
        public bool IsValidName { get; set; }
    }

    internal sealed class NlpNamePrediction
    {
        public bool IsValid { get; set; }

        public float Probability { get; set; }
    }

    internal sealed class NlpShapeCandidate
    {
        public string PageName { get; set; }
        public int ShapeId { get; set; }
        public string ShapeName { get; set; }
        public string Label { get; set; }
        public string ShapeType { get; set; }
    }

    internal sealed class NlpCheckResult
    {
        public NlpShapeCandidate Candidate { get; set; }
        public bool IsValid { get; set; }
        public float Probability { get; set; }
        public string Suggestions { get; set; }
    }
}
