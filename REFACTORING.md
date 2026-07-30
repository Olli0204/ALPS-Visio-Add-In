# Refactoring Status

The behavior-preserving source refactoring is complete. Remaining work listed in
`ALPS_Visio_AddIn-rewrite/TODO.md` concerns ontology coverage, automated Windows
tests, packaging, or changes that first require Visio characterization tests.
`FUNCTIONAL_SCOPE.md` records the user-facing comparison with `main`.

## Goals and Boundaries

The refactoring preserves observable Visio behavior while reducing global state
and separating VSTO, ontology, export, and COM concerns:

- **Add-in UI:** `ThisAddIn` and `ALPSRibbon` own lifecycle and user interaction.
- **Importing:** resource resolution, parser composition, workflow, and UI errors
  are separate.
- **Visio infrastructure:** ShapeSheet access, stencils, page creation, shape
  positioning, routing, and Auto-Arrange have focused implementations.
- **Model adapters:** `OWLShapes` maps PASS/ALPS elements to shared exporters.
- **Interactive editing:** `Snapping` owns document-scoped page controllers;
  Model Explorer tree construction is separate from its WPF event handling.

`VisioHelper` remains as a source-compatible facade so existing model adapters
and legacy callers do not need a risky all-at-once migration.
`ALPSConstants`, `ALPSGlobalFunctions`, and historically misspelled public type
names are retained as compatibility surfaces; new code uses `Constants` and
focused infrastructure classes.

## Completed Work

- [x] Extract ShapeSheet access and invariant formula formatting.
- [x] Move My Shapes discovery and stencil version selection behind the
  source-compatible `ShapeFinder` facade.
- [x] Extract stencil lookup, caching, opening, and master placement.
- [x] Isolate stencil-window restoration and drawing activation from the
  `VisioHelper` compatibility facade.
- [x] Extract SID/SBD page creation and collision-safe naming.
- [x] Extract routing, Auto-Arrange, and normalized shape positioning.
- [x] Separate Visio graph placement from deterministic connector rebinding.
- [x] Split OWL resource resolution, parser composition, import workflow, and UI.
- [x] Replace Ribbon singleton use with lifecycle-owned importer composition.
- [x] Consolidate shared behavior export and SID page sizing.
- [x] Replace ID-keyed/static layout dictionaries with weak, object-keyed state.
- [x] Separate fallback graph ranking and weak layout state from the
  `VisioLayout` facade.
- [x] Separate fallback connector-port planning and model-bounds resolution
  from the `VisioLayout` and `VisioHelper` facades.
- [x] Scope snapping controllers to the active document and release their COM
  event handlers on refresh, document switches, and shutdown.
- [x] Extract Model Explorer tree construction and priority normalization from
  the WPF code-behind while retaining its legacy XAML path.

- [x] Promote active snapping and compatibility code out of `_old`.

## Deliberately Deferred

Automated unit tests, Windows CI, and reproducible VSTO packaging require a
Windows runner with Microsoft Visio/VSTO and are not added in this source-only
pass. The Model Explorer remains under `_old/UI` because moving its WPF sources
triggers MC1000 in the legacy WinFX markup compiler. Its established namespace,
XAML path, and event behavior remain unchanged. Further behavioral changes to
snapping and WPF interaction require characterization tests. A lifecycle-adapter
extraction was rolled back after the Model Explorer stopped responding in Visio.

## Quality Gates

Keep public entry points compatible and `ALPS_Visio_AddIn-rewrite.csproj` valid.
Run `git diff --check` for every change. Before merging, build Debug and Release
on Windows and complete every item in `MANUAL_TESTING.md`.
