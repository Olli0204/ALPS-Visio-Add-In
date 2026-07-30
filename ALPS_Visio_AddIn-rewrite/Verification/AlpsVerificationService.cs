using System;
using System.Collections.Generic;
using System.Linq;

namespace ALPS_Visio_AddIn_rewrite.Verification
{
    /// <summary>
    /// Ports the implemented SID checks from ALPS-Verification-Thesis into
    /// a deterministic, UI-independent verification service.
    /// </summary>
    internal sealed class AlpsVerificationService
    {
        public VerificationReport Verify(string specificationFile,
            string implementationFile)
        {
            OwlVerificationModel specification =
                OwlVerificationModel.Load(specificationFile);
            OwlVerificationModel implementation =
                OwlVerificationModel.Load(implementationFile);
            VerificationReport report = new VerificationReport(
                specificationFile, implementationFile);

            CheckModelContents(
                specification, implementation, report);
            CheckImplementationCoverage(
                OwlElementKind.Subject, "SID", "SUBJECT",
                specification, implementation, report);
            CheckImplementationCoverage(
                OwlElementKind.MessageExchange, "SID", "MESSAGE",
                specification, implementation, report);
            CheckImplementationCoverage(
                OwlElementKind.CommunicationAct, "SID", "COMMUNICATION_ACT",
                specification, implementation, report);
            CheckFullySpecifiedSubjects(
                specification, implementation, report);
            CheckCommunicationRestrictions(
                specification, implementation, report);

            report.Findings.Add(new VerificationFinding
            {
                Severity = VerificationSeverity.Information,
                Category = "Scope",
                Code = "SCOPE-001",
                Message = "Die integrierte Prüfung deckt den im Thesis-Repo "
                    + "implementierten SID-Teil ab. SBD-Ablauflogik, "
                    + "Präzedenz-/Trigger-Transitionen, abstrakte "
                    + "Kommunikationskanäle, Finalized-Message-Semantik und "
                    + "Subjekt-Multiplizitäten sind noch nicht vollständig "
                    + "formal verifiziert."
            });

            return report;
        }

        private static void CheckModelContents(
            OwlVerificationModel specification,
            OwlVerificationModel implementation,
            VerificationReport report)
        {
            report.CheckedRuleCount += 2;
            if (!specification.Elements.Any(element =>
                element.Kind != OwlElementKind.Other))
            {
                report.Findings.Add(new VerificationFinding
                {
                    Severity = VerificationSeverity.Error,
                    Category = "Input",
                    Code = "MODEL-001",
                    Message = "Die Spezifikationsdatei enthält keine "
                        + "erkannten ALPS/PASS-Elemente."
                });
            }

            if (!implementation.Elements.Any(element =>
                element.Kind != OwlElementKind.Other))
            {
                report.Findings.Add(new VerificationFinding
                {
                    Severity = VerificationSeverity.Error,
                    Category = "Input",
                    Code = "MODEL-002",
                    Message = "Die Implementierungsdatei enthält keine "
                        + "erkannten ALPS/PASS-Elemente."
                });
            }
        }

        private static void CheckImplementationCoverage(
            OwlElementKind kind, string category, string codePrefix,
            OwlVerificationModel specification,
            OwlVerificationModel implementation,
            VerificationReport report)
        {
            IList<OwlVerificationElement> implementationElements =
                implementation.ElementsOfKind(kind).ToList();

            foreach (OwlVerificationElement specified in
                specification.ElementsOfKind(kind))
            {
                report.CheckedRuleCount++;
                IList<OwlVerificationElement> matches =
                    FindImplementations(specified,
                        implementationElements);

                if (matches.Count == 0)
                {
                    report.Findings.Add(new VerificationFinding
                    {
                        Severity = VerificationSeverity.Error,
                        Category = category,
                        Code = codePrefix + "-001",
                        SpecificationElement = specified.DisplayName,
                        Message = "Für das spezifizierte Element wurde keine "
                            + "Implementierung mit einer passenden "
                            + "ALPS-implements-Referenz gefunden."
                    });
                }
                else if (matches.Count > 1)
                {
                    report.Findings.Add(new VerificationFinding
                    {
                        Severity = VerificationSeverity.Warning,
                        Category = category,
                        Code = codePrefix + "-002",
                        SpecificationElement = specified.DisplayName,
                        ImplementationElement = string.Join(", ",
                            matches.Select(match => match.DisplayName)),
                        Message = "Das spezifizierte Element wird mehrfach "
                            + "implementiert. Prüfen Sie, ob diese "
                            + "Verfeinerung beabsichtigt ist."
                    });
                }
            }
        }

        private static void CheckFullySpecifiedSubjects(
            OwlVerificationModel specification,
            OwlVerificationModel implementation,
            VerificationReport report)
        {
            IList<OwlVerificationElement> implementationSubjects =
                implementation.ElementsOfKind(
                    OwlElementKind.Subject).ToList();

            foreach (OwlVerificationElement specified in
                specification.ElementsOfKind(
                    OwlElementKind.Subject)
                .Where(subject =>
                    subject.HasType("FullySpecifiedSubject")))
            {
                foreach (OwlVerificationElement implemented in
                    FindImplementations(specified,
                        implementationSubjects))
                {
                    report.CheckedRuleCount++;
                    if (implemented.HasType("FullySpecifiedSubject"))
                        continue;

                    report.Findings.Add(new VerificationFinding
                    {
                        Severity = VerificationSeverity.Error,
                        Category = "SID",
                        Code = "SUBJECT-TYPE-001",
                        SpecificationElement = specified.DisplayName,
                        ImplementationElement = implemented.DisplayName,
                        Message = "Ein FullySpecifiedSubject muss durch ein "
                            + "FullySpecifiedSubject implementiert werden."
                    });
                }
            }
        }

        private static void CheckCommunicationRestrictions(
            OwlVerificationModel specification,
            OwlVerificationModel implementation,
            VerificationReport report)
        {
            IList<OwlVerificationElement> implementationMessages =
                implementation.ElementsOfKind(
                    OwlElementKind.MessageExchange).ToList();

            foreach (OwlVerificationElement restriction in
                specification.ElementsOfKind(
                    OwlElementKind.CommunicationRestriction))
            {
                report.CheckedRuleCount++;
                IList<string> correspondentReferences =
                    GetCorrespondentReferences(restriction);
                IList<OwlVerificationElement> correspondents =
                    correspondentReferences
                        .Select(specification.Resolve)
                        .Where(element => element != null)
                        .Distinct()
                        .ToList();

                if (correspondents.Count != 2)
                {
                    report.Findings.Add(new VerificationFinding
                    {
                        Severity = VerificationSeverity.Warning,
                        Category = "SID",
                        Code = "RESTRICTION-001",
                        SpecificationElement = restriction.DisplayName,
                        Message = "Die Kommunikationsbeschränkung besitzt "
                            + "nicht genau zwei auflösbare Korrespondenten "
                            + "und konnte daher nicht vollständig geprüft "
                            + "werden."
                    });
                    continue;
                }

                foreach (OwlVerificationElement message in
                    implementationMessages)
                {
                    OwlVerificationElement sender =
                        ResolveFirst(implementation,
                            message.SenderReferences);
                    OwlVerificationElement receiver =
                        ResolveFirst(implementation,
                            message.ReceiverReferences);

                    if (sender == null || receiver == null)
                    {
                        report.Findings.Add(new VerificationFinding
                        {
                            Severity = VerificationSeverity.Warning,
                            Category = "SID",
                            Code = "RESTRICTION-002",
                            SpecificationElement =
                                restriction.DisplayName,
                            ImplementationElement = message.DisplayName,
                            Message = "Sender oder Empfänger des "
                                + "implementierenden Nachrichtenaustauschs "
                                + "konnte nicht aufgelöst werden."
                        });
                        continue;
                    }

                    bool forward =
                        ImplementsOrMatches(sender, correspondents[0])
                        && ImplementsOrMatches(receiver,
                            correspondents[1]);
                    bool reverse =
                        ImplementsOrMatches(sender, correspondents[1])
                        && ImplementsOrMatches(receiver,
                            correspondents[0]);

                    if (!forward && !reverse)
                        continue;

                    report.Findings.Add(new VerificationFinding
                    {
                        Severity = VerificationSeverity.Error,
                        Category = "SID",
                        Code = "RESTRICTION-003",
                        SpecificationElement = restriction.DisplayName,
                        ImplementationElement = message.DisplayName,
                        Message = "Der Nachrichtenaustausch verletzt die "
                            + "spezifizierte Kommunikationsbeschränkung "
                            + "zwischen den beiden Subjekten."
                    });
                }
            }
        }

        private static IList<OwlVerificationElement>
            FindImplementations(OwlVerificationElement specified,
                IEnumerable<OwlVerificationElement>
                    implementationElements)
        {
            return implementationElements
                .Where(implemented =>
                    implemented.ImplementedReferences.Any(reference =>
                        OwlVerificationModel.ReferenceMatches(
                            reference, specified)))
                .ToList();
        }

        private static IList<string> GetCorrespondentReferences(
            OwlVerificationElement restriction)
        {
            List<string> references =
                restriction.CorrespondentReferences.ToList();
            foreach (string sender in restriction.SenderReferences)
            {
                if (!references.Contains(sender))
                    references.Add(sender);
            }
            foreach (string receiver in restriction.ReceiverReferences)
            {
                if (!references.Contains(receiver))
                    references.Add(receiver);
            }
            return references;
        }

        private static OwlVerificationElement ResolveFirst(
            OwlVerificationModel model, IEnumerable<string> references)
        {
            foreach (string reference in references)
            {
                OwlVerificationElement resolved = model.Resolve(reference);
                if (resolved != null)
                    return resolved;
            }
            return null;
        }

        private static bool ImplementsOrMatches(
            OwlVerificationElement implementation,
            OwlVerificationElement specification)
        {
            if (OwlVerificationModel.ReferenceMatches(
                implementation.NodeId, specification))
            {
                return true;
            }

            return implementation.ImplementedReferences.Any(reference =>
                OwlVerificationModel.ReferenceMatches(
                    reference, specification));
        }
    }
}
