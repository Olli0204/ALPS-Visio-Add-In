using alps.net.api.StandardPASS;
using System.Runtime.CompilerServices;

namespace ALPS_Visio_AddIn_rewrite.OWLShapes.Layout
{
    /// <summary>
    /// Stores generated geometry without extending the lifetime of model elements.
    /// A reset starts an isolated layout pass for the next imported model or behavior.
    /// </summary>
    internal sealed class FallbackLayoutState
    {
        private ConditionalWeakTable<IPASSProcessModelElement, LayoutBounds> generatedBounds =
            new ConditionalWeakTable<IPASSProcessModelElement, LayoutBounds>();
        private ConditionalWeakTable<ITransition, TransitionPorts> portsByTransition =
            new ConditionalWeakTable<ITransition, TransitionPorts>();

        public void ResetBounds()
        {
            generatedBounds =
                new ConditionalWeakTable<IPASSProcessModelElement, LayoutBounds>();
        }

        public void ResetTransitionPorts()
        {
            portsByTransition = new ConditionalWeakTable<ITransition, TransitionPorts>();
        }

        public bool TryGetBounds(IPASSProcessModelElement element, out LayoutBounds bounds)
        {
            bounds = null;
            return element != null && generatedBounds.TryGetValue(element, out bounds);
        }

        public bool HasBounds(IPASSProcessModelElement element)
        {
            return element != null
                && generatedBounds.TryGetValue(element, out LayoutBounds ignored);
        }

        public void SetBounds(IPASSProcessModelElement element, double x, double y,
            double width, double height)
        {
            generatedBounds.Remove(element);
            generatedBounds.Add(element, new LayoutBounds(x, y, width, height));
        }

        public bool TryGetTransitionPorts(ITransition transition,
            out TransitionPorts ports)
        {
            ports = null;
            return transition != null
                && portsByTransition.TryGetValue(transition, out ports);
        }

        public TransitionPorts GetOrCreateTransitionPorts(ITransition transition)
        {
            return portsByTransition.GetValue(
                transition, ignored => new TransitionPorts());
        }

        internal sealed class LayoutBounds
        {
            public LayoutBounds(double x, double y, double width, double height)
            {
                X = x;
                Y = y;
                Width = width;
                Height = height;
            }

            public double X { get; private set; }
            public double Y { get; private set; }
            public double Width { get; private set; }
            public double Height { get; private set; }
        }

        internal sealed class TransitionPorts
        {
            public double SourceY { get; set; } = 0.5;
            public double TargetY { get; set; } = 0.5;
            public bool IsFeedback { get; set; }
            public bool UseFallbackRouting { get; set; }
        }
    }
}
