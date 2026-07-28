# Refactoring Status

## Goals and Boundaries

The refactoring preserves observable Visio behavior while reducing global state
and separating VSTO, ontology, export, and COM concerns:

- **Add-in UI:** `ThisAddIn` and `ALPSRibbon` own lifecycle and user interaction.
- **Importing:** resource resolution, parser composition, workflow, and UI errors
  are separate.
- **Visio infrastructure:** ShapeSheet access, stencils, page creation, shape
  positioning, routing, and Auto-Arrange have focused implementations.
- **Model adapters:** `OWLShapes` maps PASS/ALPS elements to shared exporters.
- **Interactive editing:** `Snapping` and the Model Explorer in `_old/UI` remain directly
  connected to `ThisAddIn` until characterization tests allow behavioral
  decomposition.

`VisioHelper` remains as a source-compatible facade so existing model adapters
and legacy callers do not need a risky all-at-once migration.

## Completed Work

- [x] Extract ShapeSheet access and invariant formula formatting.
- [x] Extract stencil lookup, caching, opening, and master placement.
- [x] Extract SID/SBD page creation and collision-safe naming.
- [x] Extract routing, Auto-Arrange, and normalized shape positioning.
- [x] Split OWL resource resolution, parser composition, import workflow, and UI.
- [x] Replace Ribbon singleton use with lifecycle-owned importer composition.
- [x] Consolidate shared behavior export and SID page sizing.
- [x] Replace ID-keyed/static layout dictionaries with weak, object-keyed state.

- [x] Promote active snapping and compatibility code out of `_old`.

## Deliberately Deferred

Automated unit tests, Windows CI, and reproducible VSTO packaging require a
Windows runner with Microsoft Visio/VSTO and are not added in this source-only
pass. Snapping now has a first-class module directory. The Model Explorer remains
under `_old/UI` because moving its WPF sources triggers MC1000 in the legacy WinFX
markup compiler. Both retain their established namespaces and behavior. Internal
decomposition is deferred until characterization tests exist. A lifecycle-adapter
extraction was rolled back after the Model Explorer stopped responding in Visio.

## Quality Gates

Keep public entry points compatible and `ALPS_Visio_AddIn-rewrite.csproj` valid.
Run `git diff --check` for every change. Before merging, build Debug and Release
on Windows and complete every item in `MANUAL_TESTING.md`.
