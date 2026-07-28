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

## 2. Lifecycle and Stencils

- Create a blank document, open another document, switch between windows, and
  close one document. No exception dialog should appear.
- Click **Open ALPS/PASS Stencils**. SID and SBD stencils should open with their
  VBA macros enabled.
- Import an OWL file and click the stencil button again. Both stencils must
  remain interactive; their macro-driven actions must still respond.
- Open **Show layer Explorer**. Switch pages/documents and confirm the tree
  refreshes, and verify that refresh, navigation, and editing remain responsive.

## 3. OWL Import

Import `docs/[Test]_Vacation_Request_2D.owl`.

- A SID and its linked SBD pages are created with unique names.
- Shapes use the supplied coordinates and retain labels, IDs, comments, and
  hyperlinks.
- Message exchanges and state transitions connect the correct endpoints.

Import `docs/[Test]_Vacation_Request.owl`.

- Missing coordinates trigger deterministic fallback placement.
- SID/SBD page sizes remain readable; shapes do not overlap unexpectedly.
- Forward, feedback, parallel, and self-loop transitions route visibly.
- Repeating either import creates unique page names and does not reuse routing
  state from the previous model.

## 4. Editing and Layout

- On both SID and SBD pages, run **Auto-Arrange → Top-down** and **Left-right**.
  Connectors must remain glued to their original source and target.
- Move subjects and states; verify snapping and connector updates.
- Create or edit SID/SBD pages with the stencil macros and confirm the layer
  explorer reflects changes.
- Verify SID-to-SBD navigation and any `extends`/background-page relationships.

Record failures with the input OWL file, page name, action, exception text, and
a before/after screenshot.
