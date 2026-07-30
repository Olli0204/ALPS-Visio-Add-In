using PassBpmnConverter.Bpmn.BpmnDI;
using PassBpmnConverter.Bpmn.DC;
using PassBpmnConverter.Bpmn.DI;

namespace PassBpmnConverter.Bpmn;

public class BpmnDiagramGenerator
{
    private const double GridColumnSize = 150;
    private const double GridRowSize = 140;
    private const double ParticipantHorizontalPadding = 65;
    private const double ParticipantVerticalPadding = 55;
    private const double ParticipantSpacing = 120;
    private const double EdgeClearance = 25;
    private const double BackEdgeClearance = 50;
    private const double BackEdgeSpacing = 30;

    private readonly Dictionary<IFlowElementsContainer, Grid?> _grids = new Dictionary<IFlowElementsContainer, Grid?>();

    public static void GenerateDiagram(IBpmnModel bpmnModel)
    {
        new BpmnDiagramGenerator().GenerateModelDiagram(bpmnModel);
    }

    private void GenerateModelDiagram(IBpmnModel bpmnModel)
    {
        if (bpmnModel == null || bpmnModel.Definitions == null)
            return;

        ICollaboration? collaboration = bpmnModel.Definitions.RootElements.OfType<ICollaboration>().FirstOrDefault();

        if (collaboration == null)
            return;

        GenerateLayout(collaboration);

        IBpmnPlane bpmnPlane = GenerateDiagram(collaboration);

        IBpmnDiagram bpmnDiagram = new BpmnDiagram()
        {
            Id = BpmnUtility.GenerateUniqueIdentifier(),
            BpmnPlane = bpmnPlane
        };

        bpmnModel.Definitions.Diagrams.Clear();
        bpmnModel.Definitions.Diagrams.Add(bpmnDiagram);
    }

    private void GenerateLayout(ICollaboration collaboration)
    {
        foreach (IParticipant participant in collaboration.Participants)
        {
            if (participant.ProcessRef != null)
            {
                _grids[participant.ProcessRef] = GenerateLayout(participant.ProcessRef);
            }
        }
    }

    private Grid? GenerateLayout(IFlowElementsContainer flowElementsContainer)
    {
        Grid grid = new Grid();

        Queue<(IFlowNode current, IFlowNode? previous)> queue = new Queue<(IFlowNode current, IFlowNode? previous)>();

        Dictionary<IFlowNode, List<IBoundaryEvent>> boundaryEventsByHost =
            new Dictionary<IFlowNode, List<IBoundaryEvent>>();

        List<IFlowNode> flowNodes = flowElementsContainer.FlowElements
            .OfType<IFlowNode>()
            .Where(flowNode => flowNode is not IBoundaryEvent)
            .ToList();

        // make sure event sub processes are added below base process
        IEnumerable<IFlowNode> initialFlowNodes = flowNodes
            .Where(flowNode => flowNode.Incoming.Count == 0)
            .OrderBy(flowNode => flowNode is ISubProcess);

        foreach (IFlowNode flowNode in initialFlowNodes)
        {
            queue.Enqueue((flowNode, null));
        }

        foreach (IBoundaryEvent boundaryEvent in flowElementsContainer.FlowElements.OfType<IBoundaryEvent>())
        {
            if (boundaryEventsByHost.TryGetValue(
                boundaryEvent.AttachedToRef,
                out List<IBoundaryEvent>? boundaryEvents))
            {
                boundaryEvents.Add(boundaryEvent);
            }
            else
            {
                boundaryEventsByHost[boundaryEvent.AttachedToRef] =
                    new List<IBoundaryEvent> { boundaryEvent };
            }
        }

        while (queue.Count > 0 || flowNodes.Any(flowNode => !grid.Contains(flowNode)))
        {
            // A valid process normally has a start node. This fallback also lays
            // out disconnected components and processes that only consist of a cycle.
            if (queue.Count == 0)
            {
                IFlowNode nextComponent = flowNodes.First(flowNode => !grid.Contains(flowNode));
                queue.Enqueue((nextComponent, null));
            }

            var elements = queue.Dequeue();
            (IFlowNode current, IFlowNode? previous) = elements;

            if (grid.Contains(current))
                continue;

            if (current is not IBoundaryEvent && previous != null && grid.Contains(previous))
            {
                grid.AddAfter(current, previous);
            }
            else
            {
                grid.Add(current);
            }

            if (current is IFlowElementsContainer subFlowElementsContainer)
            {
                if (!_grids.ContainsKey(subFlowElementsContainer))
                {
                    Grid? subGrid = GenerateLayout(subFlowElementsContainer);
                    _grids[subFlowElementsContainer] = subGrid;
                }
            }

            if (boundaryEventsByHost.TryGetValue(
                current,
                out List<IBoundaryEvent>? boundaryEventsForCurrent))
            {
                foreach (IBoundaryEvent boundaryEvent in boundaryEventsForCurrent)
                {
                    queue.Enqueue((boundaryEvent, current));
                }
            }

            IEnumerable<ISequenceFlow> outgoing = current.Outgoing;
            if (current is IExclusiveGateway exclusiveGateway
                && exclusiveGateway.Default != null)
            {
                // The conditioned path is the visual main path. A default branch
                // becomes a secondary row instead of displacing the main flow.
                outgoing = outgoing.OrderBy(
                    sequenceFlow => ReferenceEquals(
                        sequenceFlow,
                        exclusiveGateway.Default));
            }

            foreach (ISequenceFlow sequenceFlow in outgoing)
            {
                queue.Enqueue((sequenceFlow.TargetRef, current));
            }
        }

        return grid;
    }

    private IBpmnPlane GenerateDiagram(ICollaboration collaboration)
    {
        IBpmnPlane bpmnPlane = new BpmnPlane()
        {
            Id = GenerateDiagramIdentifier(collaboration),
            BpmnElement = collaboration,
        };

        double currentY = 0;

        for (int participantIndex = 0;
             participantIndex < collaboration.Participants.Count;
             participantIndex++)
        {
            IParticipant participant =
                collaboration.Participants[participantIndex];
            IBounds participantBounds;

            if (participant.ProcessRef != null && _grids.TryGetValue(participant.ProcessRef, out Grid? grid) && grid != null)
            {
                bool routeBackEdgesAbove =
                    participantIndex
                    < collaboration.Participants.Count / 2d;
                List<IDiagramElement> diagramElements = GenerateDiagram(
                    grid,
                    routeBackEdgesAbove);
                if (diagramElements.OfType<IBpmnShape>().Any())
                {
                    IBounds contentBounds = GetDiagramBounds(diagramElements);
                    double offsetX =
                        -contentBounds.X + ParticipantHorizontalPadding;
                    double offsetY =
                        currentY - contentBounds.Y + ParticipantVerticalPadding;

                    participantBounds = new Bounds()
                    {
                        X = 0,
                        Y = currentY,
                        Width =
                            contentBounds.Width
                            + ParticipantHorizontalPadding * 2,
                        Height =
                            contentBounds.Height
                            + ParticipantVerticalPadding * 2
                    };

                    // Shift contained nodes and all routed edge waypoints together.
                    foreach (IDiagramElement diagramElement in diagramElements)
                    {
                        if (diagramElement is IBpmnShape bpmnShape)
                        {
                            IBounds bounds = bpmnShape.Bounds;
                            bounds.X += offsetX;
                            bounds.Y += offsetY;
                            bpmnShape.Bounds = bounds;
                        }
                        if (diagramElement is IBpmnEdge bpmnEdge)
                        {
                            foreach (IPoint waypoint in bpmnEdge.Waypoints)
                            {
                                waypoint.X += offsetX;
                                waypoint.Y += offsetY;
                            }

                            if (bpmnEdge.BpmnLabel?.Bounds != null)
                            {
                                bpmnEdge.BpmnLabel.Bounds.X += offsetX;
                                bpmnEdge.BpmnLabel.Bounds.Y += offsetY;
                            }
                        }
                        bpmnPlane.DiagramElements.Add(diagramElement);
                    }
                }
                else
                {
                    (int width, int height) = GetDefaultShapeSize(participant);
                    participantBounds = new Bounds()
                    {
                        X = 0,
                        Y = currentY,
                        Width = width,
                        Height = height
                    };
                }
            }
            else
            {
                (int width, int height) = GetDefaultShapeSize(participant);
                participantBounds = new Bounds()
                {
                    X = 0,
                    Y = currentY,
                    Width = width,
                    Height = height
                };
            }

            IBpmnShape participantBpmnShape = new BpmnShape()
            {
                Id = GenerateDiagramIdentifier(participant),
                BpmnElement = participant,
                Bounds = participantBounds,
                IsHorizontal = true,
            };
            bpmnPlane.DiagramElements.Add(participantBpmnShape);

            currentY = participantBounds.Y + participantBounds.Height + ParticipantSpacing;
        }

        // Stacked pools should form one clean column. Individual processes can
        // still use different heights, but unequal widths make message flows and
        // the overall collaboration unnecessarily jagged.
        List<IBpmnShape> participantShapes = bpmnPlane.DiagramElements
            .OfType<IBpmnShape>()
            .Where(shape => shape.BpmnElement is IParticipant)
            .ToList();
        if (participantShapes.Count > 0)
        {
            double commonParticipantWidth =
                participantShapes.Max(shape => shape.Bounds.Width);
            foreach (IBpmnShape participantShape in participantShapes)
            {
                participantShape.Bounds.Width = commonParticipantWidth;
            }
        }

        GenerateMessageFlowDiagramElements(
            collaboration,
            bpmnPlane);

        return bpmnPlane;
    }

    private static IBounds GetDiagramBounds(
        IEnumerable<IDiagramElement> diagramElements)
    {
        double minX = double.PositiveInfinity;
        double minY = double.PositiveInfinity;
        double maxX = double.NegativeInfinity;
        double maxY = double.NegativeInfinity;

        foreach (IDiagramElement diagramElement in diagramElements)
        {
            if (diagramElement is IBpmnShape shape)
            {
                minX = Math.Min(minX, shape.Bounds.X);
                minY = Math.Min(minY, shape.Bounds.Y);
                maxX = Math.Max(
                    maxX,
                    shape.Bounds.X + shape.Bounds.Width);
                maxY = Math.Max(
                    maxY,
                    shape.Bounds.Y + shape.Bounds.Height);
            }
            else if (diagramElement is IBpmnEdge edge)
            {
                foreach (IPoint waypoint in edge.Waypoints)
                {
                    minX = Math.Min(minX, waypoint.X);
                    minY = Math.Min(minY, waypoint.Y);
                    maxX = Math.Max(maxX, waypoint.X);
                    maxY = Math.Max(maxY, waypoint.Y);
                }
            }
        }

        return new Bounds()
        {
            X = minX,
            Y = minY,
            Width = maxX - minX,
            Height = maxY - minY
        };
    }

    private static void GenerateMessageFlowDiagramElements(
        ICollaboration collaboration,
        IBpmnPlane bpmnPlane)
    {
        IList<IBpmnShape> shapes =
            bpmnPlane.DiagramElements.OfType<IBpmnShape>().ToList();
        IList<IBpmnShape> flowNodeShapes = shapes
            .Where(shape => shape.BpmnElement is IFlowNode)
            .ToList();
        IList<IBpmnShape> participantShapes = shapes
            .Where(shape => shape.BpmnElement is IParticipant)
            .ToList();
        IList<IBpmnEdge> routedMessageEdges =
            new List<IBpmnEdge>();

        int messageFlowCount = collaboration.MessageFlows.Count;
        double messageLaneSpacing = messageFlowCount <= 1
            ? 0
            : Math.Min(24, 60d / (messageFlowCount - 1));
        int messageFlowIndex = 0;

        foreach (IMessageFlow messageFlow in collaboration.MessageFlows)
        {
            int currentMessageFlowIndex = messageFlowIndex++;
            IBpmnShape? source = shapes.FirstOrDefault(
                shape => ReferenceEquals(
                    shape.BpmnElement,
                    messageFlow.SourceRef));
            IBpmnShape? target = shapes.FirstOrDefault(
                shape => ReferenceEquals(
                    shape.BpmnElement,
                    messageFlow.TargetRef));

            if (source == null || target == null)
            {
                Console.WriteLine(
                    $"Warning: Cannot create a diagram edge for "
                    + $"{nameof(IMessageFlow)} {messageFlow.Id} because "
                    + "its source or target shape is missing.");
                continue;
            }

            double corridorY = GetMessageCorridorY(
                source,
                target,
                participantShapes)
                + (
                      currentMessageFlowIndex
                      - (messageFlowCount - 1) / 2d)
                * messageLaneSpacing;

            IBpmnEdge edge = CreateRoutedEdge(
                messageFlow,
                source,
                target,
                CreateMessageFlowRoute(
                    source,
                    target,
                    flowNodeShapes,
                    routedMessageEdges,
                    corridorY));
            edge.MessageVisibleKind = MessageVisibleKind.initiating;

            bpmnPlane.DiagramElements.Add(edge);
            routedMessageEdges.Add(edge);
        }
    }

    private List<IDiagramElement> GenerateDiagram(
        Grid grid,
        bool routeBackEdgesAbove)
    {
        List<IBpmnShape> bpmnShapes = new List<IBpmnShape>();
        List<IDiagramElement> diagramElements = new List<IDiagramElement>();

        List<(IFlowNode flowNode, int row, int col)> elements = grid.GetAllElementsWithPosition();

        // flow nodes must be added before sequence flows
        foreach ((IFlowNode flowNode, int row, int col) in elements)
        {
            IBpmnShape shape = CreateShape(flowNode, row, col);

            if (flowNode is IBoundaryEvent boundaryEvent)
            {
                IBounds attachedToBounds = bpmnShapes.First(bpmnShape => bpmnShape.BpmnElement == boundaryEvent.AttachedToRef).Bounds;
                shape.Bounds = new Bounds()
                {
                    X = attachedToBounds.X + attachedToBounds.Width / 2 - shape.Bounds.Width / 2,
                    Y = attachedToBounds.Y + attachedToBounds.Height - shape.Bounds.Height / 2,
                    Width = shape.Bounds.Width,
                    Height = shape.Bounds.Height
                };
            }

            if (flowNode is IFlowElementsContainer flowElementsContainer)
            {
                if (_grids.TryGetValue(flowElementsContainer, out Grid? subGrid) && subGrid != null)
                {
                    List<IDiagramElement> subElements = GenerateDiagram(
                        subGrid,
                        routeBackEdgesAbove);
                    diagramElements.AddRange(subElements);
                }
            }

            bpmnShapes.Add(shape);
            diagramElements.Add(shape);
        }

        IList<ISequenceFlow> sequenceFlows = elements
            .Select(element => element.flowNode)
            .SelectMany(flowNode => flowNode.Outgoing)
            .ToList();
        IEnumerable<ISequenceFlow> orderedSequenceFlows =
            sequenceFlows
                .Where(sequenceFlow => !IsLongBackEdge(
                    sequenceFlow,
                    bpmnShapes))
                .Concat(
                    sequenceFlows
                        .Where(sequenceFlow => IsLongBackEdge(
                            sequenceFlow,
                            bpmnShapes))
                        .OrderByDescending(sequenceFlow =>
                            GetBoundsCenter(
                                bpmnShapes.First(shape =>
                                    ReferenceEquals(
                                        shape.BpmnElement,
                                        sequenceFlow.SourceRef))
                                    .Bounds).Y));

        int backEdgeIndex = 0;
        foreach (ISequenceFlow sequenceFlow in orderedSequenceFlows)
        {
            IBpmnShape? source = bpmnShapes.FirstOrDefault(
                bpmnShape => ReferenceEquals(
                    bpmnShape.BpmnElement,
                    sequenceFlow.SourceRef));
            IBpmnShape? target = bpmnShapes.FirstOrDefault(
                bpmnShape => ReferenceEquals(
                    bpmnShape.BpmnElement,
                    sequenceFlow.TargetRef));

            if (source == null || target == null)
            {
                Console.WriteLine($"Warning: Cannot find {nameof(IDiagramElement)} for source or target of {nameof(ISequenceFlow)}. {nameof(ISequenceFlow)} will not have a {nameof(IDiagramElement)}.");
                continue;
            }

            IBpmnEdge edge = CreateSequenceFlowEdge(
                sequenceFlow,
                source,
                target,
                bpmnShapes,
                routeBackEdgesAbove,
                ref backEdgeIndex);
            diagramElements.Add(edge);
        }

        return diagramElements;
    }

    private static bool IsLongBackEdge(
        ISequenceFlow sequenceFlow,
        IEnumerable<IBpmnShape> shapes)
    {
        IBpmnShape? source = shapes.FirstOrDefault(shape =>
            ReferenceEquals(
                shape.BpmnElement,
                sequenceFlow.SourceRef));
        IBpmnShape? target = shapes.FirstOrDefault(shape =>
            ReferenceEquals(
                shape.BpmnElement,
                sequenceFlow.TargetRef));
        return source != null
               && target != null
               && GetBoundsCenter(source.Bounds).X
               - GetBoundsCenter(target.Bounds).X
               > GridColumnSize * 1.25;
    }

    private static string GenerateDiagramIdentifier(IBaseElement baseElement)
    {
        return baseElement.Id + "_di";
    }

    private static IBpmnShape CreateShape(IBaseElement baseElement, int row, int col)
    {
        (int width, int height) = GetDefaultShapeSize(baseElement);

        IBounds bounds = new Bounds()
        {
            X = (col * GridColumnSize) - width / 2,
            Y = (row * GridRowSize) - height / 2,
            Width = width,
            Height = height
        };

        IBpmnShape bpmnShape = new BpmnShape()
        {
            Id = GenerateDiagramIdentifier(baseElement),
            BpmnElement = baseElement,
            Bounds = bounds,
        };

        return bpmnShape;
    }

    private static (int width, int height) GetDefaultShapeSize(IBaseElement element)
    {
        if (element is IParticipant)
        {
            return (400, 100);
        }

        if (element is ISubProcess)
        {
            return (100, 80);
        }

        if (element is ITask)
        {
            return (100, 80);
        }

        if (element is IGateway)
        {
            return (50, 50);
        }

        if (element is IEvent)
        {
            return (36, 36);
        }

        return (100, 80);
    }

    private static IBpmnEdge CreateSequenceFlowEdge(
        ISequenceFlow sequenceFlow,
        IBpmnShape source,
        IBpmnShape target,
        IList<IBpmnShape> shapes,
        bool routeBackEdgesAbove,
        ref int backEdgeIndex)
    {
        IPoint sourceCenter = GetBoundsCenter(source.Bounds);
        IPoint targetCenter = GetBoundsCenter(target.Bounds);
        List<IPoint> waypoints;
        double? feedbackCorridorY = null;

        if (ReferenceEquals(source, target))
        {
            double loopX = source.Bounds.X
                           + source.Bounds.Width
                           + EdgeClearance;
            double loopY = source.Bounds.Y - EdgeClearance;
            waypoints = new List<IPoint>()
            {
                GetRightDock(source),
                new Point { X = loopX, Y = sourceCenter.Y },
                new Point { X = loopX, Y = loopY },
                new Point { X = sourceCenter.X, Y = loopY },
                GetTopDock(target)
            };
        }
        else if (NearlyEqual(sourceCenter.X, targetCenter.X))
        {
            if (IsVerticalSegmentClear(
                    sourceCenter.X,
                    sourceCenter.Y,
                    targetCenter.Y,
                    shapes,
                    source,
                    target))
            {
                waypoints = targetCenter.Y > sourceCenter.Y
                    ? new List<IPoint>()
                    {
                        GetBottomDock(source),
                        GetTopDock(target)
                    }
                    : new List<IPoint>()
                    {
                        GetTopDock(source),
                        GetBottomDock(target)
                    };
            }
            else
            {
                double laneX = Math.Max(
                                   source.Bounds.X + source.Bounds.Width,
                                   target.Bounds.X + target.Bounds.Width)
                               + EdgeClearance;
                waypoints = new List<IPoint>()
                {
                    GetRightDock(source),
                    new Point { X = laneX, Y = sourceCenter.Y },
                    new Point { X = laneX, Y = targetCenter.Y },
                    GetRightDock(target)
                };
            }
        }
        else if (targetCenter.X > sourceCenter.X)
        {
            IPoint sourceDock = GetRightDock(source);
            IPoint targetDock = GetLeftDock(target);

            if (NearlyEqual(sourceCenter.Y, targetCenter.Y))
            {
                waypoints = new List<IPoint>()
                {
                    sourceDock,
                    targetDock
                };
            }
            else
            {
                double laneX =
                    (sourceDock.X + targetDock.X) / 2;
                waypoints = new List<IPoint>()
                {
                    sourceDock,
                    new Point { X = laneX, Y = sourceCenter.Y },
                    new Point { X = laneX, Y = targetCenter.Y },
                    targetDock
                };
            }
        }
        else if (sourceCenter.X - targetCenter.X
                 <= GridColumnSize * 1.25)
        {
            // A short return to the preceding column is clearer as a compact
            // dogleg than as a loop around the complete process.
            IPoint sourceDock = GetLeftDock(source);
            IPoint targetDock = GetRightDock(target);
            double laneX = (sourceDock.X + targetDock.X) / 2;
            waypoints = NearlyEqual(sourceCenter.Y, targetCenter.Y)
                ? new List<IPoint>()
                {
                    sourceDock,
                    targetDock
                }
                : new List<IPoint>()
                {
                    sourceDock,
                    new Point { X = laneX, Y = sourceCenter.Y },
                    new Point { X = laneX, Y = targetCenter.Y },
                    targetDock
                };
        }
        else
        {
            int currentBackEdgeIndex = backEdgeIndex++;
            double corridorY = routeBackEdgesAbove
                ? shapes.Min(shape => shape.Bounds.Y)
                  - BackEdgeClearance
                  - currentBackEdgeIndex * BackEdgeSpacing
                : shapes.Max(
                      shape => shape.Bounds.Y
                               + shape.Bounds.Height)
                  + BackEdgeClearance
                  + currentBackEdgeIndex * BackEdgeSpacing;
            feedbackCorridorY = corridorY;
            if (currentBackEdgeIndex == 0)
            {
                waypoints = CreateVerticalCorridorRoute(
                    source,
                    target,
                    shapes,
                    corridorY);
            }
            else
            {
                double sourceLaneX = shapes.Max(
                    shape => shape.Bounds.X
                             + shape.Bounds.Width)
                    + EdgeClearance
                    * (currentBackEdgeIndex + 1);
                waypoints = CreateVerticalCorridorRoute(
                    CreateOuterVerticalEscape(
                        source,
                        corridorY,
                        shapes,
                        sourceLaneX),
                    CreateVerticalEscape(
                        target,
                        corridorY,
                        shapes));
            }
        }

        IBpmnEdge edge = CreateRoutedEdge(
            sequenceFlow,
            source,
            target,
            waypoints);
        if (feedbackCorridorY.HasValue)
        {
            edge.BpmnLabel = CreateFeedbackFlowLabel(
                sequenceFlow.Name,
                edge.Waypoints,
                feedbackCorridorY.Value);
        }

        return edge;
    }

    private static IBpmnEdge CreateRoutedEdge(
        IBaseElement bpmnElement,
        IBpmnShape source,
        IBpmnShape target,
        IEnumerable<IPoint> waypoints)
    {
        return new BpmnEdge()
        {
            Id = GenerateDiagramIdentifier(bpmnElement),
            BpmnElement = bpmnElement,
            SourceElement = source,
            TargetElement = target,
            Waypoints = SimplifyWaypoints(waypoints)
        };
    }

    private static List<IPoint> CreateVerticalCorridorRoute(
        IBpmnShape source,
        IBpmnShape target,
        IList<IBpmnShape> shapes,
        double corridorY)
    {
        List<IPoint> sourceEscape = CreateVerticalEscape(
            source,
            corridorY,
            shapes);
        List<IPoint> targetEscape = CreateVerticalEscape(
            target,
            corridorY,
            shapes);

        return CreateVerticalCorridorRoute(
            sourceEscape,
            targetEscape);
    }

    private static List<IPoint> CreateVerticalCorridorRoute(
        IList<IPoint> sourceEscape,
        IList<IPoint> targetEscape)
    {
        List<IPoint> waypoints = new List<IPoint>(sourceEscape);
        for (int index = targetEscape.Count - 1; index >= 0; index--)
        {
            waypoints.Add(targetEscape[index]);
        }

        return SimplifyWaypoints(waypoints);
    }

    private static List<IPoint> CreateOuterVerticalEscape(
        IBpmnShape shape,
        double corridorY,
        IList<IBpmnShape> shapes,
        double laneX)
    {
        IPoint center = GetBoundsCenter(shape.Bounds);
        bool exitsRight = laneX >= center.X;
        IPoint dock = exitsRight
            ? GetRightDock(shape)
            : GetLeftDock(shape);

        if (IsHorizontalSegmentClear(
                dock.Y,
                dock.X,
                laneX,
                shapes,
                shape)
            && IsVerticalSegmentClear(
                laneX,
                dock.Y,
                corridorY,
                shapes,
                shape,
                null))
        {
            return new List<IPoint>()
            {
                dock,
                new Point { X = laneX, Y = dock.Y },
                new Point { X = laneX, Y = corridorY }
            };
        }

        return CreateVerticalEscape(
            shape,
            corridorY,
            shapes);
    }

    private static List<IPoint> CreateMessageFlowRoute(
        IBpmnShape source,
        IBpmnShape target,
        IList<IBpmnShape> shapes,
        IList<IBpmnEdge> routedMessageEdges,
        double corridorY)
    {
        List<IPoint> defaultSourceEscape =
            CreateVerticalEscape(
                source,
                corridorY,
                shapes);
        List<IPoint> defaultTargetEscape =
            CreateVerticalEscape(
                target,
                corridorY,
                shapes);
        List<IPoint> defaultRoute =
            CreateVerticalCorridorRoute(
                defaultSourceEscape,
                defaultTargetEscape);
        if (routedMessageEdges.Count == 0
            || CountEdgeConflicts(
                defaultRoute,
                routedMessageEdges) == 0)
        {
            return defaultRoute;
        }

        IList<List<IPoint>> targetCandidates =
            CreateMessageEscapeCandidates(
                target,
                corridorY,
                shapes);

        List<IPoint> bestRoute = defaultRoute;
        int bestConflictCount = CountEdgeConflicts(
            bestRoute,
            routedMessageEdges);
        double bestLength = GetRouteLength(bestRoute);

        foreach (List<IPoint> targetEscape in targetCandidates)
        {
            List<IPoint> route =
                CreateVerticalCorridorRoute(
                    defaultSourceEscape,
                    targetEscape);
            int conflictCount = CountEdgeConflicts(
                route,
                routedMessageEdges);
            double length = GetRouteLength(route);
            if (conflictCount < bestConflictCount
                || (
                    conflictCount == bestConflictCount
                    && length < bestLength))
            {
                bestRoute = route;
                bestConflictCount = conflictCount;
                bestLength = length;
            }
        }

        return bestRoute;
    }

    private static IList<List<IPoint>> CreateMessageEscapeCandidates(
        IBpmnShape shape,
        double corridorY,
        IList<IBpmnShape> shapes)
    {
        List<List<IPoint>> candidates =
            new List<List<IPoint>>();
        IPoint center = GetBoundsCenter(shape.Bounds);
        bool exitsDown = corridorY >= center.Y;
        IPoint verticalDock = exitsDown
            ? GetBottomDock(shape)
            : GetTopDock(shape);

        if (IsVerticalSegmentClear(
                center.X,
                verticalDock.Y,
                corridorY,
                shapes,
                shape,
                null))
        {
            candidates.Add(
                new List<IPoint>()
                {
                    verticalDock,
                    new Point
                    {
                        X = center.X,
                        Y = corridorY
                    }
                });
        }

        const int maximumLaneMultiplier = 4;
        for (int multiplier = 1;
             multiplier <= maximumLaneMultiplier;
             multiplier++)
        {
            double distance = EdgeClearance * multiplier;
            AddMessageSideEscapeCandidate(
                candidates,
                shape,
                corridorY,
                shapes,
                shape.Bounds.X
                + shape.Bounds.Width
                + distance,
                exitsDown,
                true);
            AddMessageSideEscapeCandidate(
                candidates,
                shape,
                corridorY,
                shapes,
                shape.Bounds.X - distance,
                exitsDown,
                false);
        }

        if (candidates.Count == 0)
        {
            candidates.Add(
                CreateVerticalEscape(
                    shape,
                    corridorY,
                    shapes));
        }

        return candidates;
    }

    private static void AddMessageSideEscapeCandidate(
        ICollection<List<IPoint>> candidates,
        IBpmnShape shape,
        double corridorY,
        IList<IBpmnShape> shapes,
        double laneX,
        bool exitsDown,
        bool exitsRight)
    {
        IPoint dock = GetMessageSideDock(
            shape,
            exitsRight,
            exitsDown);
        if (!IsHorizontalSegmentClear(
                dock.Y,
                dock.X,
                laneX,
                shapes,
                shape)
            || !IsVerticalSegmentClear(
                laneX,
                dock.Y,
                corridorY,
                shapes,
                shape,
                null))
        {
            return;
        }

        candidates.Add(
            new List<IPoint>()
            {
                dock,
                new Point { X = laneX, Y = dock.Y },
                new Point { X = laneX, Y = corridorY }
            });
    }

    private static IPoint GetMessageSideDock(
        IBpmnShape shape,
        bool exitsRight,
        bool exitsDown)
    {
        double inset = Math.Min(
            10,
            shape.Bounds.Height / 4);
        return new Point()
        {
            X = exitsRight
                ? shape.Bounds.X + shape.Bounds.Width
                : shape.Bounds.X,
            Y = exitsDown
                ? shape.Bounds.Y + shape.Bounds.Height - inset
                : shape.Bounds.Y + inset
        };
    }

    private static int CountEdgeConflicts(
        IList<IPoint> route,
        IEnumerable<IBpmnEdge> routedEdges)
    {
        int conflictCount = 0;
        for (int routeIndex = 1;
             routeIndex < route.Count;
             routeIndex++)
        {
            IPoint routeStart = route[routeIndex - 1];
            IPoint routeEnd = route[routeIndex];

            foreach (IBpmnEdge routedEdge in routedEdges)
            {
                for (int edgeIndex = 1;
                     edgeIndex < routedEdge.Waypoints.Count;
                     edgeIndex++)
                {
                    if (SegmentsConflict(
                            routeStart,
                            routeEnd,
                            routedEdge.Waypoints[edgeIndex - 1],
                            routedEdge.Waypoints[edgeIndex]))
                    {
                        conflictCount++;
                    }
                }
            }
        }

        return conflictCount;
    }

    private static bool SegmentsConflict(
        IPoint firstStart,
        IPoint firstEnd,
        IPoint secondStart,
        IPoint secondEnd)
    {
        bool firstHorizontal =
            NearlyEqual(firstStart.Y, firstEnd.Y);
        bool secondHorizontal =
            NearlyEqual(secondStart.Y, secondEnd.Y);

        if (firstHorizontal == secondHorizontal)
            return false;

        IPoint horizontalStart = firstHorizontal
            ? firstStart
            : secondStart;
        IPoint horizontalEnd = firstHorizontal
            ? firstEnd
            : secondEnd;
        IPoint verticalStart = firstHorizontal
            ? secondStart
            : firstStart;
        IPoint verticalEnd = firstHorizontal
            ? secondEnd
            : firstEnd;
        return verticalStart.X
                   >= Math.Min(
                       horizontalStart.X,
                       horizontalEnd.X)
               && verticalStart.X
                   <= Math.Max(
                       horizontalStart.X,
                       horizontalEnd.X)
               && horizontalStart.Y
                   >= Math.Min(
                       verticalStart.Y,
                       verticalEnd.Y)
               && horizontalStart.Y
                   <= Math.Max(
                       verticalStart.Y,
                       verticalEnd.Y);
    }

    private static double GetRouteLength(
        IList<IPoint> route)
    {
        double length = 0;
        for (int index = 1; index < route.Count; index++)
        {
            length += Math.Abs(route[index].X - route[index - 1].X)
                      + Math.Abs(
                          route[index].Y
                          - route[index - 1].Y);
        }

        return length;
    }

    private static List<IPoint> CreateVerticalEscape(
        IBpmnShape shape,
        double corridorY,
        IList<IBpmnShape> shapes)
    {
        IPoint center = GetBoundsCenter(shape.Bounds);
        bool exitsDown = corridorY >= center.Y;
        IPoint verticalDock = exitsDown
            ? GetBottomDock(shape)
            : GetTopDock(shape);

        if (IsVerticalSegmentClear(
                center.X,
                verticalDock.Y,
                corridorY,
                shapes,
                shape,
                null))
        {
            return new List<IPoint>()
            {
                verticalDock,
                new Point { X = center.X, Y = corridorY }
            };
        }

        for (int multiplier = 1; multiplier <= 4; multiplier++)
        {
            double distance = EdgeClearance * multiplier;
            (double laneX, IPoint dock)[] candidates =
            {
                (
                    shape.Bounds.X + shape.Bounds.Width + distance,
                    GetRightDock(shape)),
                (
                    shape.Bounds.X - distance,
                    GetLeftDock(shape))
            };

            foreach ((double laneX, IPoint dock) in candidates)
            {
                if (!IsHorizontalSegmentClear(
                        dock.Y,
                        dock.X,
                        laneX,
                        shapes,
                        shape))
                {
                    continue;
                }

                if (IsVerticalSegmentClear(
                        laneX,
                        dock.Y,
                        corridorY,
                        shapes,
                        shape,
                        null))
                {
                    return new List<IPoint>()
                    {
                        dock,
                        new Point { X = laneX, Y = dock.Y },
                        new Point { X = laneX, Y = corridorY }
                    };
                }
            }
        }

        double fallbackLaneX =
            shape.Bounds.X + shape.Bounds.Width + EdgeClearance * 5;
        IPoint fallbackDock = GetRightDock(shape);
        return new List<IPoint>()
        {
            fallbackDock,
            new Point { X = fallbackLaneX, Y = fallbackDock.Y },
            new Point { X = fallbackLaneX, Y = corridorY }
        };
    }

    private static bool IsVerticalSegmentClear(
        double x,
        double fromY,
        double toY,
        IEnumerable<IBpmnShape> shapes,
        IBpmnShape ignoredShape,
        IBpmnShape? secondIgnoredShape)
    {
        const double obstaclePadding = 10;
        double minY = Math.Min(fromY, toY);
        double maxY = Math.Max(fromY, toY);

        foreach (IBpmnShape shape in shapes)
        {
            if (ReferenceEquals(shape, ignoredShape)
                || ReferenceEquals(shape, secondIgnoredShape)
                || shape.BpmnElement is not IFlowNode)
            {
                continue;
            }

            IBounds bounds = shape.Bounds;
            bool intersectsX =
                x >= bounds.X - obstaclePadding
                && x <= bounds.X + bounds.Width + obstaclePadding;
            bool intersectsY =
                maxY >= bounds.Y - obstaclePadding
                && minY <= bounds.Y + bounds.Height + obstaclePadding;
            if (intersectsX && intersectsY)
                return false;
        }

        return true;
    }

    private static bool IsHorizontalSegmentClear(
        double y,
        double fromX,
        double toX,
        IEnumerable<IBpmnShape> shapes,
        IBpmnShape ignoredShape)
    {
        const double obstaclePadding = 10;
        double minX = Math.Min(fromX, toX);
        double maxX = Math.Max(fromX, toX);

        foreach (IBpmnShape shape in shapes)
        {
            if (ReferenceEquals(shape, ignoredShape)
                || shape.BpmnElement is not IFlowNode)
            {
                continue;
            }

            IBounds bounds = shape.Bounds;
            bool intersectsX =
                maxX >= bounds.X - obstaclePadding
                && minX <= bounds.X + bounds.Width + obstaclePadding;
            bool intersectsY =
                y >= bounds.Y - obstaclePadding
                && y <= bounds.Y + bounds.Height + obstaclePadding;
            if (intersectsX && intersectsY)
                return false;
        }

        return true;
    }

    private static double GetMessageCorridorY(
        IBpmnShape source,
        IBpmnShape target,
        IList<IBpmnShape> participantShapes)
    {
        IBpmnShape? sourceParticipant = FindContainingParticipant(
            source,
            participantShapes);
        IBpmnShape? targetParticipant = FindContainingParticipant(
            target,
            participantShapes);

        if (sourceParticipant == null
            || targetParticipant == null
            || ReferenceEquals(sourceParticipant, targetParticipant))
        {
            return (
                       GetBoundsCenter(source.Bounds).Y
                       + GetBoundsCenter(target.Bounds).Y)
                   / 2;
        }

        if (sourceParticipant.Bounds.Y < targetParticipant.Bounds.Y)
        {
            return (
                       sourceParticipant.Bounds.Y
                       + sourceParticipant.Bounds.Height
                       + targetParticipant.Bounds.Y)
                   / 2;
        }

        return (
                   targetParticipant.Bounds.Y
                   + targetParticipant.Bounds.Height
                   + sourceParticipant.Bounds.Y)
               / 2;
    }

    private static IBpmnLabel? CreateFeedbackFlowLabel(
        string? name,
        IEnumerable<IPoint> waypoints,
        double corridorY)
    {
        if (string.IsNullOrWhiteSpace(name))
            return null;

        List<IPoint> points = waypoints.ToList();
        double longestSegmentLength = 0;
        double centerX = points.Average(waypoint => waypoint.X);
        for (int index = 1; index < points.Count; index++)
        {
            IPoint previous = points[index - 1];
            IPoint current = points[index];
            if (!NearlyEqual(previous.Y, corridorY)
                || !NearlyEqual(current.Y, corridorY))
            {
                continue;
            }

            double segmentLength = Math.Abs(current.X - previous.X);
            if (segmentLength > longestSegmentLength)
            {
                longestSegmentLength = segmentLength;
                centerX = (previous.X + current.X) / 2;
            }
        }

        double width = Math.Min(
            180,
            Math.Max(90, name.Length * 7));
        const double height = 20;

        return new BpmnLabel()
        {
            Bounds = new Bounds()
            {
                X = centerX - width / 2,
                Y = corridorY - height - 6,
                Width = width,
                Height = height
            }
        };
    }

    private static IBpmnShape? FindContainingParticipant(
        IBpmnShape elementShape,
        IEnumerable<IBpmnShape> participantShapes)
    {
        IPoint center = GetBoundsCenter(elementShape.Bounds);
        return participantShapes.FirstOrDefault(participantShape =>
            center.X >= participantShape.Bounds.X
            && center.X <= participantShape.Bounds.X
                + participantShape.Bounds.Width
            && center.Y >= participantShape.Bounds.Y
            && center.Y <= participantShape.Bounds.Y
                + participantShape.Bounds.Height);
    }

    private static List<IPoint> SimplifyWaypoints(
        IEnumerable<IPoint> waypoints)
    {
        List<IPoint> simplified = new List<IPoint>();

        foreach (IPoint waypoint in waypoints)
        {
            if (simplified.Count > 0)
            {
                IPoint previous = simplified[simplified.Count - 1];
                if (NearlyEqual(previous.X, waypoint.X)
                    && NearlyEqual(previous.Y, waypoint.Y))
                {
                    continue;
                }
            }

            simplified.Add(
                new Point
                {
                    X = waypoint.X,
                    Y = waypoint.Y
                });

            while (simplified.Count >= 3)
            {
                IPoint first = simplified[simplified.Count - 3];
                IPoint middle = simplified[simplified.Count - 2];
                IPoint last = simplified[simplified.Count - 1];
                bool vertical =
                    NearlyEqual(first.X, middle.X)
                    && NearlyEqual(middle.X, last.X);
                bool horizontal =
                    NearlyEqual(first.Y, middle.Y)
                    && NearlyEqual(middle.Y, last.Y);
                if (!vertical && !horizontal)
                    break;

                simplified.RemoveAt(simplified.Count - 2);
            }
        }

        return simplified;
    }

    private static bool NearlyEqual(double first, double second)
    {
        return Math.Abs(first - second) < 0.001;
    }

    private static IPoint GetBoundsCenter(IBounds bounds)
    {
        return new Point()
        {
            X = bounds.X + bounds.Width / 2,
            Y = bounds.Y + bounds.Height / 2
        };
    }

    private static IPoint GetLeftDock(IBpmnShape shape)
    {
        return new Point()
        {
            X = shape.Bounds.X,
            Y = shape.Bounds.Y + shape.Bounds.Height / 2
        };
    }

    private static IPoint GetRightDock(IBpmnShape shape)
    {
        return new Point()
        {
            X = shape.Bounds.X + shape.Bounds.Width,
            Y = shape.Bounds.Y + shape.Bounds.Height / 2
        };
    }

    private static IPoint GetTopDock(IBpmnShape shape)
    {
        return new Point()
        {
            X = shape.Bounds.X + shape.Bounds.Width / 2,
            Y = shape.Bounds.Y
        };
    }

    private static IPoint GetBottomDock(IBpmnShape shape)
    {
        return new Point()
        {
            X = shape.Bounds.X + shape.Bounds.Width / 2,
            Y = shape.Bounds.Y + shape.Bounds.Height
        };
    }
}
