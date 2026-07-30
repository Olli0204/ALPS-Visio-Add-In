using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using VDS.RDF;
using VDS.RDF.Parsing;

namespace ALPS_Visio_AddIn_rewrite.Verification
{
    internal enum OwlElementKind
    {
        Other,
        Subject,
        MessageExchange,
        CommunicationAct,
        State,
        Transition,
        CommunicationRestriction
    }

    internal sealed class OwlVerificationElement
    {
        public OwlVerificationElement(string nodeId)
        {
            NodeId = nodeId;
            Types = new HashSet<string>(StringComparer.Ordinal);
            ImplementedReferences = new List<string>();
            SenderReferences = new List<string>();
            ReceiverReferences = new List<string>();
            CorrespondentReferences = new List<string>();
        }

        public string NodeId { get; }

        public string ComponentId { get; set; }

        public string Label { get; set; }

        public ISet<string> Types { get; }

        public IList<string> ImplementedReferences { get; }

        public IList<string> SenderReferences { get; }

        public IList<string> ReceiverReferences { get; }

        public IList<string> CorrespondentReferences { get; }

        public OwlElementKind Kind
        {
            get
            {
                if (HasType("CommunicationRestriction"))
                    return OwlElementKind.CommunicationRestriction;
                if (Types.Any(IsMessageExchangeType))
                    return OwlElementKind.MessageExchange;
                if (Types.Any(IsSubjectType))
                    return OwlElementKind.Subject;
                if (Types.Contains("SendFunction")
                    || Types.Contains("ReceiveFunction")
                    || Types.Contains("CommunicationAct"))
                {
                    return OwlElementKind.CommunicationAct;
                }
                if (Types.Any(IsTransitionType))
                    return OwlElementKind.Transition;
                if (Types.Any(IsStateType))
                    return OwlElementKind.State;
                return OwlElementKind.Other;
            }
        }

        public string DisplayName
        {
            get
            {
                string id = string.IsNullOrWhiteSpace(ComponentId)
                    ? OwlVerificationModel.LocalName(NodeId)
                    : ComponentId;
                if (string.IsNullOrWhiteSpace(Label)
                    || string.Equals(Label, id,
                        StringComparison.Ordinal))
                {
                    return id;
                }

                return Label + " (" + id + ")";
            }
        }

        public bool HasType(string type)
        {
            return Types.Contains(type);
        }

        private static bool IsMessageExchangeType(string type)
        {
            return type.EndsWith("MessageExchange",
                       StringComparison.Ordinal)
                || type.EndsWith("MessageConnector",
                       StringComparison.Ordinal);
        }

        private static bool IsSubjectType(string type)
        {
            if (!type.EndsWith("Subject", StringComparison.Ordinal))
                return false;

            return !string.Equals(type, "SubjectBehavior",
                       StringComparison.Ordinal)
                && !string.Equals(type, "SubjectBaseBehavior",
                       StringComparison.Ordinal)
                && !string.Equals(type, "SubjectDataDefinition",
                       StringComparison.Ordinal);
        }

        private static bool IsTransitionType(string type)
        {
            return type.EndsWith("Transition", StringComparison.Ordinal)
                && !type.EndsWith("TransitionCondition",
                    StringComparison.Ordinal);
        }

        private static bool IsStateType(string type)
        {
            return type.EndsWith("State", StringComparison.Ordinal)
                && !string.Equals(type, "InitialStateOfBehavior",
                    StringComparison.Ordinal)
                && !string.Equals(type, "EndState",
                    StringComparison.Ordinal);
        }
    }

    internal sealed class OwlVerificationModel
    {
        private readonly IDictionary<string, OwlVerificationElement>
            elementsByNode;

        private OwlVerificationModel(
            IDictionary<string, OwlVerificationElement> elementsByNode)
        {
            this.elementsByNode = elementsByNode;
        }

        public IEnumerable<OwlVerificationElement> Elements =>
            elementsByNode.Values;

        public IEnumerable<OwlVerificationElement> ElementsOfKind(
            OwlElementKind kind)
        {
            return Elements.Where(element => element.Kind == kind);
        }

        public OwlVerificationElement Resolve(string reference)
        {
            if (string.IsNullOrWhiteSpace(reference))
                return null;

            if (elementsByNode.TryGetValue(reference, out
                OwlVerificationElement direct))
            {
                return direct;
            }

            return Elements.FirstOrDefault(element =>
                ReferenceMatches(reference, element));
        }

        public static OwlVerificationModel Load(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException(
                    "An OWL file name is required.", nameof(fileName));
            if (!File.Exists(fileName))
                throw new FileNotFoundException(
                    "The OWL model could not be found.", fileName);

            Graph graph = new Graph();
            FileLoader.Load(graph, fileName);

            Dictionary<string, OwlVerificationElement> elements =
                new Dictionary<string, OwlVerificationElement>(
                    StringComparer.Ordinal);

            foreach (Triple triple in graph.Triples)
            {
                if (!string.Equals(LocalName(NodeValue(triple.Predicate)),
                    "type", StringComparison.Ordinal))
                {
                    continue;
                }

                string nodeId = NodeValue(triple.Subject);
                if (!elements.TryGetValue(nodeId,
                    out OwlVerificationElement element))
                {
                    element = new OwlVerificationElement(nodeId);
                    elements.Add(nodeId, element);
                }

                element.Types.Add(LocalName(NodeValue(triple.Object)));
            }

            foreach (Triple triple in graph.Triples)
            {
                string nodeId = NodeValue(triple.Subject);
                if (!elements.TryGetValue(nodeId,
                    out OwlVerificationElement element))
                {
                    continue;
                }

                string predicate = LocalName(
                    NodeValue(triple.Predicate));
                string value = NodeValue(triple.Object);

                switch (predicate)
                {
                    case "hasModelComponentID":
                        element.ComponentId = value;
                        break;
                    case "hasModelComponentLabel":
                    case "label":
                        if (string.IsNullOrWhiteSpace(element.Label))
                            element.Label = value;
                        break;
                    case "implements":
                        AddDistinct(element.ImplementedReferences, value);
                        break;
                    case "hasSender":
                    case "sender":
                        AddDistinct(element.SenderReferences, value);
                        break;
                    case "hasReceiver":
                    case "receiver":
                        AddDistinct(element.ReceiverReferences, value);
                        break;
                    case "hasCorrespondent":
                    case "correspondent":
                        AddDistinct(element.CorrespondentReferences, value);
                        break;
                }
            }

            return new OwlVerificationModel(elements);
        }

        public static bool ReferenceMatches(string reference,
            OwlVerificationElement element)
        {
            if (element == null || string.IsNullOrWhiteSpace(reference))
                return false;

            string normalizedReference = Normalize(reference);
            string normalizedNode = Normalize(element.NodeId);
            string normalizedId = Normalize(element.ComponentId);

            return string.Equals(normalizedReference, normalizedNode,
                       StringComparison.Ordinal)
                || (!string.IsNullOrEmpty(normalizedId)
                    && string.Equals(normalizedReference, normalizedId,
                        StringComparison.Ordinal))
                || string.Equals(LocalName(normalizedReference),
                    LocalName(normalizedNode), StringComparison.Ordinal)
                || (!string.IsNullOrEmpty(normalizedId)
                    && string.Equals(LocalName(normalizedReference),
                        normalizedId, StringComparison.Ordinal));
        }

        public static string LocalName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string normalized = Normalize(value);
            int hash = normalized.LastIndexOf('#');
            int slash = normalized.LastIndexOf('/');
            int separator = Math.Max(hash, slash);
            return separator >= 0 && separator + 1 < normalized.Length
                ? normalized.Substring(separator + 1)
                : normalized;
        }

        private static string NodeValue(INode node)
        {
            IUriNode uriNode = node as IUriNode;
            if (uriNode != null)
                return uriNode.Uri.OriginalString;

            ILiteralNode literalNode = node as ILiteralNode;
            if (literalNode != null)
                return literalNode.Value;

            return node?.ToString() ?? string.Empty;
        }

        private static string Normalize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string trimmed = value.Trim().Trim('<', '>');
            try
            {
                return Uri.UnescapeDataString(trimmed);
            }
            catch (UriFormatException)
            {
                return trimmed;
            }
        }

        private static void AddDistinct(IList<string> values,
            string value)
        {
            if (!string.IsNullOrWhiteSpace(value)
                && !values.Contains(value))
            {
                values.Add(value);
            }
        }
    }
}
