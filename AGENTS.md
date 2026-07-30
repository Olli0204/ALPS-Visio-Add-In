# Repository Guidelines

## Project Structure & Module Organization

`ALPS_Visio_Tools.sln` contains the .NET Framework 4.8 VSTO add-in in `ALPS_Visio_AddIn-rewrite/`. `ThisAddIn.cs` and `ALPSRibbon.cs` are lifecycle/UI entry points. Import orchestration lives in `Importing/`; adapters and export logic live in `OWLShapes/`, with fallback layout under `OWLShapes/Layout/`. Put reusable Visio COM concerns in `VisioInfrastructure/`. Keep `VisioHelper.cs` as a compatibility facade.

`Snapping/` and `_old/UI/` contain compatibility-sensitive production code
connected through `ThisAddIn`; controllers must remain document-scoped and
release COM events when refreshed. Preserve behavior unless Windows/Visio
characterization tests cover changes. The Model Explorer remains in `_old/UI/`
because moving its XAML sources breaks legacy WinFX markup compilation. Constants
and helpers live in `Compatibility/`. Ontologies and images are in `Resources/`;
the integrated label checker lives in `NlpChecking/`, including its embedded
training set under `NlpChecking/Resources/`. Documentation and sample OWL files
are in `docs/`. There is no active test project.

## Image Retention

Never delete repository images or files in the adjacent `Troubleshooting IMGs/`
folder. They support the bachelor thesis; report cleanup candidates without
changing them.

## Build, Test, and Development Commands

Development requires Windows, Visio, and Visual Studio 2022 with Office/VSTO.

```powershell
nuget restore ALPS_Visio_Tools.sln
msbuild ALPS_Visio_Tools.sln /p:Configuration=Debug
msbuild ALPS_Visio_Tools.sln /p:Configuration=Release
```

Restore `packages.config` dependencies before building. Run Debug from Visual Studio to launch or attach to Visio. Follow `MANUAL_TESTING.md` for the acceptance pass.

## Coding Style & Naming Conventions

Use four spaces and braces on separate lines. Use PascalCase for types and public members, camelCase for locals and private fields, and `I` prefixes for interfaces. Keep ontology adapters named `Visio<Type>`.

Centralize ShapeSheet names in `Constants` and access cells through `VisioInfrastructure`. Preserve .NET Framework 4.8/VSTO compatibility and existing public facade signatures. Document public APIs; comment only non-obvious Visio or ontology behavior.

Never add API keys, model credentials, or unencrypted user settings. The NLP
suggestion service is optional; its offline classifier must continue to work
without a key or network connection. Any new external transmission must be
explicitly disclosed in the UI and acceptance checklist.

## Testing Guidelines

No automated suite or coverage threshold exists. Manually verify both `docs/[Test]_Vacation_Request_2D.owl` and `docs/[Test]_Vacation_Request.owl`, SID/SBD links, connectors, routing, both Auto-Arrange directions, snapping, and the layer explorer. Add regression OWL samples under `docs/`. New tests belong in a separate `*Tests` project; name methods `Method_Scenario_ExpectedResult`.

## Commit & Pull Request Guidelines

Use Conventional Commit messages (`fix:`, `feat:`, `refactor:`, `chore:`), for example `refactor: isolate Visio stencil access`. Pull requests should describe affected model elements, manual test steps, known limitations, and linked issues. Include before/after screenshots for visual changes.

Push through the configured SSH remote
`git@github.com:Olli0204/ALPS-Visio-Add-In.git`. The available SSH key
authenticates as `Olli0204`; do not replace it with the failing HTTPS/`gh`
credential flow. Confirm the local branch and its upstream are synchronized.

Do not commit `bin/`, `obj/`, `packages/`, certificates, generated publish output, or machine-specific Visio paths.
