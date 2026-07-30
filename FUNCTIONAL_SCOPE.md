# Functional Scope Comparison

This comparison uses `origin/main` at `722efed` and the `codex-rewrite`
branch. `main` is the direct ancestor of the rewrite branch; no newer main
commit is missing.

## User-facing Ribbon Commands

| Command | main | codex-rewrite |
| --- | --- | --- |
| Open ALPS/PASS Stencils | Available | Available with drawing-window and macro handling |
| Show layer Explorer | Available | Available with document-safe refresh |
| Import OWL | Available | Available with refactored parser, multi-model export, and fallback layout |
| Convert PASS to BPMN | Separate command-line converter | Integrated OWL/RDF-to-BPMN 2.0 conversion with generated diagram layout |
| ALPS Verification | Visible debug placeholder | Integrated two-model SID verification with structured findings |
| Auto-Arrange | Not available | Top-down and left-right |
| NLP PASS Checking | Separate external add-in | Integrated local naming check with selectable suggestion providers and models |

The Ribbon retains the three groups from main: **Standard Functions**, **ALPS
Layer Editing**, and **OWL PASS Tools**. Auto-Arrange is added to OWL PASS
Tools. **NLP PASS Checking** is an additional group whose split button provides
the model check, classifier retraining, and provider settings.

**Model Conversion** is an additional group containing **Convert PASS to
BPMN**. The command selects an OWL/RDF PASS model, converts the first readable
model with the integrated converter, generates BPMN diagram coordinates, and
writes a `.bpmn` file selected by the user.

## ALPS Verification

The verification workflow was adapted from
`andikra/ALPS-Verification-Thesis`. The console prototype is not embedded as
an executable: its supported rules were ported into a UI-independent,
in-process OWL/RDF verification service that remains compatible with the
add-in's .NET Framework 4.8 host.

The Ribbon command asks for an abstract specification and an implementing
model, then checks:

- explicit `implements` coverage for SID subjects, message exchanges, and
  communication acts;
- preservation of `FullySpecifiedSubject`;
- communication restrictions in both directions; and
- ambiguous multiple implementations or unresolvable restriction endpoints.

The results dialog separates errors, warnings, and scope information and can
copy a tab-separated report. Parsing and checking are completely offline.
The original thesis model corpus is retained under
`docs/verification-thesis/`; small deterministic add-in regression fixtures
remain alongside it under `docs/`.

The source repository is explicitly a prototype. Its SBD checker is empty and
its README lists precedence/trigger transitions, abstract communication
channels, finalized messages, multiplicities, and start/end properties as
future work. The integrated result therefore reports this boundary and never
claims full formal conformance beyond the supported SID checks.

## NLP Naming Check

The branch integrates the behavior and training set from
`MatthesElstermann/NLPPASSCheckingBackup`. The classifier trains lazily from an
embedded TSV resource and checks supported PASS labels without network access.
If the user explicitly configures an active provider, only labels classified
for review are sent for optional suggestions. OpenAI, Anthropic, and UniGPT are
built in; users can add OpenAI-compatible or Anthropic-compatible endpoints.
Available models are queried from the provider and selected from a dropdown.
Provider profiles, selected models, and keys are stored together in a
Windows-DPAPI-protected configuration for the current user.

## Non-Ribbon Code Removed from main

The following sources were compiled in main but were not reachable as working
user features:

- `VerificationChecker` is an empty WinForms prototype; the Ribbon callback
  instead displayed “Hello World” and toggled VBA listeners.
- `OWLImportDialog` contains hard-coded sample model names and is never opened.
- `TemporaryModelExplorer` is superseded by the active WPF Model Explorer.
- `OLD_SiSi` is initialized at startup but has no Ribbon command or external
  source entry point.
- Old test projects contain one incomplete assertion and one test without an
  assertion; the trial project uses machine-specific paths.
- Checked-in publish output is generated deployment data, not source
  functionality.

These prototypes were intentionally not restored. The only visible main Ribbon
surface that was absent—the verification command—is now present without
reintroducing its unsafe debug-side effects.
