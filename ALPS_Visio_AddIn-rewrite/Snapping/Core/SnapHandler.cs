using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace VisioAddIn.Snapping
{
    public abstract class SnapHandler
    {
        private static readonly IEqualityComparer<Shape> ShapeComparer =
            new ShapeIdComparer();

        protected IDictionary<Shape, Shape> snappedShapes;
        private readonly ISet<Shape> shapesBeingAdjusted;
        private readonly ISet<Shape> shapesWithOpenMaintenanceDecision;
        private readonly IDictionary<int, int> lastPromptedCandidateByShapeId;

        /// <summary>
        /// Maximum center-to-center snap distance in millimeters.
        /// </summary>
        public const int SNAP_RANGE = 20;

        protected SnapHandler()
        {
            snappedShapes = CreateShapeDictionary();
            shapesBeingAdjusted = new HashSet<Shape>(ShapeComparer);
            shapesWithOpenMaintenanceDecision =
                new HashSet<Shape>(ShapeComparer);
            lastPromptedCandidateByShapeId = new Dictionary<int, int>();
        }

        public virtual void performSnap(
            Shape snappingShape, Shape backgroundReferenceShape)
        {
            if (snappingShape == null || backgroundReferenceShape == null)
                return;

            snappedShapes[snappingShape] = backgroundReferenceShape;
            lastPromptedCandidateByShapeId.Remove(GetShapeId(snappingShape));
            adjustSize(snappingShape, backgroundReferenceShape);
        }

        /// <summary>
        /// Checks whether the shape is close enough to the single nearest
        /// compatible background shape to offer a snap.
        /// </summary>
        public void checkForSnapping(Shape snappingShape)
        {
            if (snappingShape == null
                || shapesBeingAdjusted.Contains(snappingShape)
                || !isShapeSnappable(snappingShape))
            {
                return;
            }

            int snappingShapeId = GetShapeId(snappingShape);
            if (snappedShapes.TryGetValue(
                snappingShape, out Shape currentReference))
            {
                if (isLocatedClosely(snappingShape, currentReference))
                {
                    shapesWithOpenMaintenanceDecision.Remove(snappingShape);
                    // PinX and PinY events may still arrive after Visio has
                    // completed the drag. Reassert the exact overlay without
                    // showing another confirmation dialog.
                    adjustSize(snappingShape, currentReference);
                    return;
                }

                if (shapesWithOpenMaintenanceDecision.Add(snappingShape))
                    handleDistantSnappedShapes(snappingShape);

                if (!snappedShapes.TryGetValue(
                        snappingShape, out Shape maintainedReference)
                    || isLocatedClosely(
                        snappingShape, maintainedReference))
                {
                    shapesWithOpenMaintenanceDecision.Remove(snappingShape);
                }
                // A maintenance dialog may keep or remove the binding. Never
                // offer a different target in the same movement event.
                return;
            }

            shapesWithOpenMaintenanceDecision.Remove(snappingShape);

            Shape nearestCandidate = getSnappableShapesOnBackgroundPage()
                .Where(candidate => candidate != null
                    && isLocatedClosely(snappingShape, candidate))
                .OrderBy(candidate => GetDistanceSquared(
                    snappingShape, candidate))
                .ThenBy(GetShapeId)
                .FirstOrDefault();

            if (nearestCandidate == null)
            {
                lastPromptedCandidateByShapeId.Remove(snappingShapeId);
                return;
            }

            int candidateId = GetShapeId(nearestCandidate);
            if (lastPromptedCandidateByShapeId.TryGetValue(
                snappingShapeId, out int promptedCandidateId)
                && promptedCandidateId == candidateId)
            {
                return;
            }

            // CellChanged is raised independently for PinX and PinY. Remember
            // the offered pair before opening the modal dialog so the second
            // event cannot display the same question again.
            lastPromptedCandidateByShapeId[snappingShapeId] = candidateId;
            WindowSnapConfirmation confirmation =
                new WindowSnapConfirmation(
                    this, snappingShape, nearestCandidate);
            confirmation.ShowDialog();
        }

        protected abstract bool isShapeSnappable(IVShape shape);
        protected abstract void handleDistantSnappedShapes(
            Shape snappingShape);
        protected abstract IEnumerable<Shape>
            getSnappableShapesOnBackgroundPage();

        public abstract void snap(
            Shape snappingShape, string backgroundReferenceShapeName);
        public abstract void unsnap(Shape shape);

        /// <summary>
        /// Sets the referenced background page, including clearing it with
        /// null.
        /// </summary>
        protected abstract void setBackPage(DiagramPage newProperty);

        public void setBackgroundPage(DiagramPage newProperty)
        {
            foreach (Shape shape in snappedShapes.Keys.ToList())
                unsnap(shape);

            snappedShapes = CreateShapeDictionary();
            shapesWithOpenMaintenanceDecision.Clear();
            lastPromptedCandidateByShapeId.Clear();
            setBackPage(newProperty);
        }

        protected void adjustSize(
            Shape snappingShape, Shape backgroundReferenceShape)
        {
            if (snappingShape == null || backgroundReferenceShape == null
                || !shapesBeingAdjusted.Add(snappingShape))
            {
                return;
            }

            try
            {
                double x = GetMillimeters(
                    backgroundReferenceShape,
                    ALPSConstants.shapeCellShapeTransformPinX);
                double y = GetMillimeters(
                    backgroundReferenceShape,
                    ALPSConstants.shapeCellShapeTransformPinY);
                double width = GetMillimeters(
                    backgroundReferenceShape,
                    ALPSConstants.shapeCellShapeTransformWidth) + 5d;
                double height = GetMillimeters(
                    backgroundReferenceShape,
                    ALPSConstants.shapeCellShapeTransformHeight) + 5d;

                SetMillimeters(snappingShape,
                    ALPSConstants.shapeCellShapeTransformPinX, x);
                SetMillimeters(snappingShape,
                    ALPSConstants.shapeCellShapeTransformPinY, y);
                SetMillimeters(snappingShape,
                    ALPSConstants.shapeCellShapeTransformWidth, width);
                SetMillimeters(snappingShape,
                    ALPSConstants.shapeCellShapeTransformHeight, height);
            }
            finally
            {
                shapesBeingAdjusted.Remove(snappingShape);
            }
        }

        protected bool isLocatedClosely(Shape shape, Shape snapToShape)
        {
            double deltaX = GetMillimeters(shape, "PinX")
                - GetMillimeters(snapToShape, "PinX");
            double deltaY = GetMillimeters(shape, "PinY")
                - GetMillimeters(snapToShape, "PinY");
            return IsWithinSnapRange(deltaX, deltaY, SNAP_RANGE);
        }

        internal static bool IsWithinSnapRange(
            double deltaX, double deltaY, double snapRange)
        {
            if (snapRange < 0d)
                throw new ArgumentOutOfRangeException(nameof(snapRange));

            return deltaX * deltaX + deltaY * deltaY
                <= snapRange * snapRange;
        }

        public void notifyBackgroundShapeMoved(Shape snapToShape)
        {
            if (snapToShape == null) return;

            List<Shape> affectedShapes = snappedShapes
                .Where(pair => ShapeComparer.Equals(
                    pair.Value, snapToShape))
                .Select(pair => pair.Key)
                .ToList();

            foreach (Shape shape in affectedShapes)
                adjustSize(shape, snapToShape);
        }

        protected static bool AreSameShape(Shape first, Shape second)
        {
            return ShapeComparer.Equals(first, second);
        }

        protected static int GetShapeId(Shape shape)
        {
            if (shape == null) return -1;
            try
            {
                return shape.ID;
            }
            catch (COMException)
            {
                return RuntimeHelpers.GetHashCode(shape);
            }
        }

        private static IDictionary<Shape, Shape> CreateShapeDictionary()
        {
            return new Dictionary<Shape, Shape>(ShapeComparer);
        }

        private static double GetDistanceSquared(
            Shape first, Shape second)
        {
            double deltaX = GetMillimeters(first, "PinX")
                - GetMillimeters(second, "PinX");
            double deltaY = GetMillimeters(first, "PinY")
                - GetMillimeters(second, "PinY");
            return deltaX * deltaX + deltaY * deltaY;
        }

        private static double GetMillimeters(Shape shape, string cellName)
        {
            return shape.CellsU[cellName]
                .Result[VisUnitCodes.visMillimeters];
        }

        private static void SetMillimeters(
            Shape shape, string cellName, double value)
        {
            double currentValue = GetMillimeters(shape, cellName);
            if (Math.Abs(currentValue - value) <= 0.001d)
                return;

            shape.CellsU[cellName].FormulaU =
                value.ToString(CultureInfo.InvariantCulture) + " mm";
        }

        /// <summary>
        /// A handler observes one foreground and at most one background page,
        /// so the stable Visio shape ID is sufficient and avoids unreliable
        /// RCW reference equality across COM events.
        /// </summary>
        private sealed class ShapeIdComparer : IEqualityComparer<Shape>
        {
            public bool Equals(Shape first, Shape second)
            {
                if (ReferenceEquals(first, second)) return true;
                if (first == null || second == null) return false;
                return GetShapeId(first) == GetShapeId(second);
            }

            public int GetHashCode(Shape shape)
            {
                return GetShapeId(shape);
            }
        }
    }
}
