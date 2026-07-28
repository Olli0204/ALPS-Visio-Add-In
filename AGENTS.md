# Repository Guidelines

## Project Structure & Module Organization

`ALPS_Visio_Tools.sln` contains one .NET Framework 4.8 VSTO add-in under `ALPS_Visio_AddIn-rewrite/`. `ThisAddIn.cs` and `ALPSRibbon.cs` are UI/lifecycle entry points. OWL import orchestration lives in `Importing/`; ontology-to-Visio adapters and export logic live in `OWLShapes/`. Put reusable Visio COM concerns—ShapeSheet access, stencils, pages, routing, and layout application—in `VisioInfrastructure/`. `VisioHelper.cs` is a compatibility facade, not a home for new responsibilities.

`Snapping/` and `ModelExplorer/` contain active, compatibility-sensitive
production code connected directly through `ThisAddIn`; preserve behavior unless
Windows/Visio characterization tests cover the change. Transitional constants
and helpers live in `Compatibility/`. Ontologies and images are in `Resources/`;
documentation and sample OWL files are in `docs/`. There is no active test project.

## Build, Test, and Development Commands

Development requires Windows, Microsoft Visio, and Visual Studio 2022 with the Office/VSTO workload.

```powershell
nuget restore ALPS_Visio_Tools.sln
msbuild ALPS_Visio_Tools.sln /p:Configuration=Debug
msbuild ALPS_Visio_Tools.sln /p:Configuration=Release
```

Restore `packages.config` dependencies before building. Run Debug from Visual Studio to launch or attach to Visio. Follow `MANUAL_TESTING.md` for the acceptance pass.

## Coding Style & Naming Conventions

Use four spaces and braces on separate lines. Use PascalCase for types and public members, camelCase for locals and private fields, and `I` prefixes for interfaces. Keep ontology adapters named `Visio<Type>`.

Centralize ShapeSheet names in `Constants` and access cells through `VisioInfrastructure`. Preserve .NET Framework 4.8/VSTO compatibility and existing public facade signatures. Document public APIs; comment only non-obvious Visio or ontology behavior.

## Testing Guidelines

No automated suite or coverage threshold exists. Manually verify both `docs/[Test]_Vacation_Request_2D.owl` and `docs/[Test]_Vacation_Request.owl`, SID/SBD links, connectors, routing, both Auto-Arrange directions, snapping, and the layer explorer. Add regression OWL samples under `docs/`. New tests belong in a separate `*Tests` project; name methods `Method_Scenario_ExpectedResult`.

## Commit & Pull Request Guidelines

Use focused Conventional Commit-style messages (`fix:`, `feat:`, `refactor:`, `chore:`), for example `refactor: isolate Visio stencil access`. Pull requests should describe affected model elements, manual test steps, known limitations, and linked issues. Include before/after screenshots for visual changes.

Push through the configured SSH remote
`git@github.com:Olli0204/ALPS-Visio-Add-In.git`. The available SSH key
authenticates as `Olli0204`; do not replace it with the failing HTTPS/`gh`
credential flow. Confirm the local branch and its upstream are synchronized.

Do not commit `bin/`, `obj/`, `packages/`, certificates, generated publish output, or machine-specific Visio paths.
