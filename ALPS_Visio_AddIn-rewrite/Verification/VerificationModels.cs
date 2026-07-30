using System;
using System.Collections.Generic;
using System.Linq;

namespace ALPS_Visio_AddIn_rewrite.Verification
{
    internal enum VerificationSeverity
    {
        Information,
        Warning,
        Error
    }

    internal sealed class VerificationFinding
    {
        public VerificationSeverity Severity { get; set; }

        public string Category { get; set; }

        public string Code { get; set; }

        public string SpecificationElement { get; set; }

        public string ImplementationElement { get; set; }

        public string Message { get; set; }
    }

    internal sealed class VerificationReport
    {
        public VerificationReport(string specificationFile,
            string implementationFile)
        {
            SpecificationFile = specificationFile;
            ImplementationFile = implementationFile;
            Findings = new List<VerificationFinding>();
        }

        public string SpecificationFile { get; }

        public string ImplementationFile { get; }

        public IList<VerificationFinding> Findings { get; }

        public int CheckedRuleCount { get; set; }

        public int ErrorCount =>
            Findings.Count(finding =>
                finding.Severity == VerificationSeverity.Error);

        public int WarningCount =>
            Findings.Count(finding =>
                finding.Severity == VerificationSeverity.Warning);

        public bool PassedSupportedChecks => ErrorCount == 0;
    }
}
