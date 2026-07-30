# Manual Visio Acceptance Test

Run this checklist on Windows with Visual Studio 2022, Microsoft Visio, the
Office/VSTO workload, and the current SID/SBD stencils installed in **My
Shapes**.

## 1. Build

```powershell
nuget restore ALPS_Visio_Tools.sln
msbuild ALPS_Visio_Tools.sln /p:Configuration=Debug
msbuild ALPS_Visio_Tools.sln /p:Configuration=Release
```

Both builds must finish without warnings newly introduced by this refactoring.
Start the Debug configuration from Visual Studio and confirm that Visio opens
with the **ALPS/PASS ADDIN** Ribbon tab.

## 2. Ribbon

- **Standard Functions** contains **Open ALPS/PASS Stencils**.
- **ALPS Layer Editing** contains **Show layer Explorer**.
- **OWL PASS Tools** contains **Import OWL**, the **Auto-Arrange** split button,
  and **ALPS Verification**.
- Click **ALPS Verification** and confirm that a single informational message
  explains that verification is not implemented. It must not start or stop VBA.

## 3. Lifecycle and Stencils

- Create a blank document, open another document, switch between windows, and
  close one document. No exception dialog should appear.
- Import an OWL file. Exactly one macro security dialog should appear: accept it
  to enable the SID drop macros. The SBD stencil opens with VBA disabled.
- Click **Open ALPS/PASS Stencils** after the import. The disabled stencil
  instance should be reopened for interactive use and its macro-driven actions
  must respond.
- When the stencil publisher or location is not trusted, Visio can show one
  additional security prompt when the SBD stencil is reopened interactively.
- Open **Show layer Explorer**. Switch pages/documents and confirm the tree
  refreshes, and verify that refresh, navigation, and editing remain responsive.
- Trigger the explorer refresh repeatedly, then move one snappable shape. Each
  action and confirmation dialog must occur once; old document controllers must
  not keep reacting.

- The Output window must not stop in `ThisDocument.restartMarkos`; OWL import no
  longer stops and restarts the active VBA listener collection.

## 4. OWL Import

Import `docs/[Test]_Vacation_Request_2D.owl`.

- A SID and its linked SBD pages are created with unique names.
- Open **Show layer Explorer** immediately after import. The imported model,
  SID layers, and linked SBD pages must be present.
- Shapes use the supplied coordinates and retain labels, IDs, comments, and
  hyperlinks.
- Message exchanges and state transitions connect the correct endpoints.
- Every SID message appears inside its centered Message Box; no message list
  members remain at the lower-left page origin.

Import `docs/[Test]_Vacation_Request.owl`.

- Missing coordinates trigger deterministic fallback placement.
- SID/SBD page sizes remain readable; shapes do not overlap unexpectedly.
- Forward, feedback, parallel, and self-loop transitions route visibly.
- Repeating either import creates unique page names and does not reuse routing
  state from the previous model.

## 5. Editing and Layout

- On both SID and SBD pages, run **Auto-Arrange → Top-down** and **Left-right**.
  Connectors must remain glued to their original source and target.
- Move subjects and states; verify snapping and connector updates.
- Create or edit SID/SBD pages with the stencil macros and confirm the layer
  explorer reflects changes.
- Verify SID-to-SBD navigation and any `extends`/background-page relationships.

Record failures with the input OWL file, page name, action, exception text, and
a before/after screenshot.
