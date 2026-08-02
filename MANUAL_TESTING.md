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
  and **Verify ALPS Models**.
- **Model Conversion** contains **Convert PASS to BPMN**.
- **NLP PASS Checking** contains the **Check Model Naming** split button with
  **Check Model Naming**, **Retrain**, and **Provider Settings** menu entries.
- Click **Verify ALPS Models**, cancel the first file dialog, and confirm that
  no verification result or exception dialog appears. It must not start or
  stop VBA.

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

- Missing coordinates trigger the same deterministic Top-down graph layout,
  connector rebinding, feedback routing, and SID Message Box positioning as the
  Auto-Arrange command.
- Every hidden semantic SID connector remains glued to both subjects. Its
  presentation polyline must terminate exactly at the corresponding subject
  boundaries; no detached horizontal or vertical fragment may remain beside a
  Message Box after that box has moved to its final layout position.
- In Top-down layout, the Employee-to-Manager channel and its `Vacation Request`
  box share the right outer corridor; the Manager-to-Employee channel and its
  `Approval`/`Denial` box share the left outer corridor.
- Exactly one visible line represents each SID channel. The semantic stencil
  connector remains on the hidden, non-printing `ALPS Internal SID Semantics`
  layer, while its fixed presentation polyline is regenerated from the stored
  subject IDs after repeated Top-down and Left-right Auto-Arrange runs.
- Reopen a document previously arranged with an older Add-in build and run
  Auto-Arrange again. Guarded legacy group geometry, including horizontal
  cross-lines, vertical leaders, and line jumps, must no longer be visible.
- The target arrowhead of each visible SID route remains visible above the
  subject fill, while message containers and their list entries stay above the
  connector line so their labels remain unobstructed.
- For vertically aligned subjects, neither native channel may coincide with a
  left or right subject border. Each route must contain two horizontal legs and
  one vertical leg through its left or right message corridor. The MessageBox
  outline must remain visible after migrating a diagram from an older build.
- Move either subject manually and run Auto-Arrange again. The visible channel
  must be regenerated at the new subject boundaries without becoming glued to
  them or losing either horizontal corridor leg.
- Drag either SID Message Box away from its channel after Auto-Arrange. A
  straight leader line without arrowheads must follow the box immediately and
  terminate on the nearest point of that channel's corridor. Moving the box
  again must update the leader without rerouting the main channel.
- Verify both directions independently: `Vacation Request` must connect
  Employee to Manager, while the combined `Approval`/`Denial` channel must
  connect Manager back to Employee. Both visible routes must terminate at the
  expected subjects, and their arrowheads must still show these semantic
  directions after a second Top-down Auto-Arrange run.
- SID/SBD page sizes remain readable; shapes do not overlap unexpectedly.
- Forward, feedback, parallel, and self-loop transitions route visibly.
- Repeating either import creates unique page names and does not reuse routing
  state from the previous model.

## 5. Editing and Layout

### Snap handlers

- On a SID extension page, move an actor extension within 20 mm of two
  background actors. Exactly one confirmation dialog must open, and it must
  offer the geometrically nearest actor even when both shapes share the same
  X-coordinate.
- Accept the SID snap. The extension must align to the background actor and be
  5 mm wider and higher. Moving the background actor must update every
  extension snapped to it, not only the first one.
- Reject a snap and continue moving within the same candidate's range. No
  duplicate dialog may appear until the extension has left and re-entered the
  20-mm range.
- Remove the SID or SBD background-page relation and move an extension. It must
  not snap to a shape from the formerly referenced page.
- Repeat the nearest-candidate, accept, maintain, and unsnap checks for an SBD
  state extension. Choosing to maintain a distant snap must restore the exact
  overlay without opening parallel maintenance windows.

- On both SID and SBD pages, run **Auto-Arrange → Top-down** and **Left-right**.
  Connectors must remain glued to their original source and target.
- In **Left-right**, long message labels must reserve enough horizontal space
  between ranks, and the print page must switch to landscape without an
  internal portrait page break.
- On SBD pages, verify that the initial state starts the primary process axis,
  alternatives share one rank, feedback transitions use the outer routing
  corridors, and repeated Auto-Arrange runs do not keep moving the states.
- On SID pages, verify that only subjects participate in the graph layout.
  Message Boxes must stay next to the corresponding communication channel,
  their message list members must move with them and retain their list
  membership without a red warning outline, opposite directions and parallel
  Message Boxes must not overlap, and subjects must remain within one compact
  portrait (Top-down) or landscape (Left-right) page.
- Move subjects and states; verify snapping and connector updates.
- Create or edit SID/SBD pages with the stencil macros and confirm the layer
  explorer reflects changes.
- Verify SID-to-SBD navigation and any `extends`/background-page relationships.

## 6. NLP PASS Checking

- Without configuring an API key, click **Check Model Naming**. Confirm that no
  authentication prompt or network request occurs and that supported shapes
  from all pages appear in the results table.
- Verify that subjects, multi-subjects, message specifications, do/send/receive
  states, and do transitions are recognized. Confirm that page, Shape ID, type,
  label, result, and confidence are populated.
- Open **Provider Settings** and verify that **OpenAI**, **Anthropic**, and
  **UniGPT** are present and cannot be removed or renamed.
- For each built-in provider, enter a valid test key and click **Load models**.
  Confirm the authenticated `/models` request populates the dropdown, select a
  model, mark the provider active, save, reopen the dialog, and verify provider
  and model selection persist.
- Add a custom provider, choose **OpenAI-compatible** or
  **Anthropic-compatible**, enter its base URL and optional key, query its
  models, select one, and save. Verify custom providers can be removed while
  built-in providers cannot.
- Restart Visio and confirm
  `%LOCALAPPDATA%\ALPS Visio Add-In\nlp-providers.dat` does not contain provider
  keys, URLs, or model names as readable text. An existing legacy UniGPT key
  should migrate from `nlp-api-key.dat` automatically.
- With suggestions enabled, confirm that only the shape type and label of
  entries marked for review are sent to the active provider. Invalid or
  unavailable credentials must produce a per-row notice without aborting the
  remaining check.
- Remove the active provider key. Run the check again and confirm it remains
  fully functional offline and does not issue suggestion requests.
- Click **Retrain** and confirm that 680 bundled examples are reported. Run the
  check twice and verify the Ribbon and dialogs stay responsive.

## 7. ALPS Verification

Use `docs/[Test]_ALPS_Verification_Specification.owl` and
`docs/[Test]_ALPS_Verification_Implementation.owl` as the specification and
implementation respectively.

- Click **Verify ALPS Models** and select the specification first and the
  implementation second.
- Confirm that the result dialog names both files and reports the number of
  checked rules, errors, and warnings.
- Confirm that missing `implements` relationships are listed with their
  specification elements instead of being written to a console or a relative
  `log.txt`.
- Confirm that communication restrictions are checked in both directions and
  that `Forbidden direct request` is reported as a violating message exchange.
- Confirm that the scope row explains that SBD precedence/trigger semantics and
  the other thesis TODOs are not yet a complete formal proof.
- Click **Ergebnisse kopieren** and paste into a text editor. The copied report
  must contain both full file paths and one tab-separated row per finding.
- Repeat the workflow with the original thesis fixtures
  `docs/verification-thesis/AbstractModel.owl` and
  `docs/verification-thesis/ImplementingModel.owl`. Confirm that both files
  load locally and that their missing or incompatible implementation
  relationships appear as structured findings.
- Repeat with a malformed OWL file and confirm that a single error dialog is
  shown and Visio remains responsive.

## 8. PASS to BPMN Conversion

- Click **Convert PASS to BPMN** and select
  `docs/[Test]_Vacation_Request.owl`.
- Save the proposed output as a `.bpmn` file and confirm that the success
  dialog shows its complete path.
- Open the output in [bpmn.io](https://demo.bpmn.io/) or another BPMN 2.0
  viewer. It must load without an invalid-ID warning even though the source
  model name begins with `[Test]`.
- Confirm that participants, processes, flow nodes, sequence flows, and BPMN
  diagram coordinates are present. All coordinates must be finite and remain
  within the visible collaboration area; no edge may jump to an extremely
  large X or Y position.
- Confirm that the two pools are aligned to the same width and each process
  uses two compact rows. The main flow runs left-to-right; feedback flows use
  separate outer corridors above the upper pool or below the lower pool
  instead of crossing tasks and message flows.
- Confirm that sequence and message flows use horizontal/vertical segments
  only. No connector may run through an unrelated task, event, or gateway.
- Confirm that message flows do not cross one another. When several feedback
  flows return to the same gateway, the lower source must use the inner
  corridor and the upper source the outer corridor without crossing.
- Confirm that `Vacation Request`, `Approval`, and `Denial` are visible as
  message flows between the matching send and receive tasks in the two pools.
  Each message name must appear exactly once next to its envelope marker in
  the gap between the pools, without overlapping a pool border or another
  label.
- Cancel the input dialog and then the output dialog in separate runs. Neither
  cancellation may show an error.
- Select a malformed OWL/RDF file and confirm that one error dialog is shown
  and Visio remains usable.

Record failures with the input OWL file, page name, action, exception text, and
a before/after screenshot.
