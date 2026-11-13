# Process-grid layout and retained structural history

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Compose process layouts from managed native text cells with row/column operations and explicit retained history.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/ProcessFlowTools.cs:1` — ProcessGridCell, ProcessGridSpec, ProcessGridPlan, ProcessGridCodec, ProcessFlowTools, Context, Saved.

## Inputs, outputs, data structures and dependencies

Inputs: ProcessGridSpec/cell selection, structural action or geometry/colors. Outputs: native cell shapes, deck XML and grid/row/column tags. Dependencies: TextShapeElements, XmlPartUpdates and NativeFormattingTransaction.

## Internal execution and external integration

Planner creates bounded rows/columns/geometry and cell assignments. Create inserts cells, copies supported style before attaching new XML and stores deck registry. Change verifies exact ownership/coordinates/visibility, preflights geometry on temporary native duplicates, snapshots native state and applies new cells/state. Hide operations retain inactive invisible cells; move swaps columns. Resize adjusts whole grid; striping uses rollback transactions.

## Configuration, validation and failure behavior

<=20 rows/12columns,<=100 grids/deck, decoded<=1000 cells, XML<=150000 chars; lookup<=2000 slide shapes. Last row/column cannot be hidden. Cell style transfer copies supported properties rather than PickUp/Apply (which can transfer CustomerData).

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/ContextPanelModelTests.cs:7`.
- `tests/s0/ContextPropertiesOfficeTests.cs:66`.
- `tests/s0/DemoScenario.cs:144`.
- `tests/s0/ProcessFlowOfficeTests.cs:12`.
- `tests/s0/ProcessFlowTests.cs:20`.

## Known limitations and technical debt

Independent manual geometry changes must be undone before grid editing. Copied grids need fresh registry. Inactive history increases content until a separately designed compaction. Missing/foreign cells or changed history stop edits; failure restores geometry/XML/tags or explicitly requests Undo.

## Extension points

Add structural actions to pure planner before adapters; verify active/inactive history, source-style metadata isolation and save/reopen.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `ProcessGridCell` | class | `apps/powerpoint-vsto/ProcessFlowTools.cs:13` |
| `ProcessGridSpec` | class | `apps/powerpoint-vsto/ProcessFlowTools.cs:14` |
| `Copy` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:19` |
| `ProcessGridPlan` | class | `apps/powerpoint-vsto/ProcessFlowTools.cs:21` |
| `Validate` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:23` |
| `Kind` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:29` |
| `Bounds` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:30` |
| `Create` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:31` |
| `Change` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:35` |
| `Resize` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:44` |
| `Stripe` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:48` |
| `ProcessGridCodec` | class | `apps/powerpoint-vsto/ProcessFlowTools.cs:50` |
| `Encode` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:53` |
| `F` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:54` |
| `I` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:55` |
| `N` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:56` |
| `Fields` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:57` |
| `Decode` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:58` |
| `ProcessFlowTools` | class | `apps/powerpoint-vsto/ProcessFlowTools.cs:63` |
| `Context` | class | `apps/powerpoint-vsto/ProcessFlowTools.cs:66` |
| `Saved` | class | `apps/powerpoint-vsto/ProcessFlowTools.cs:67` |
| `Active` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:68` |
| `TextPart` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:69` |
| `Read` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:70` |
| `Place` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:74` |
| `Tags` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:75` |
| `CopyColor` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:76` |
| `CopyCellStyle` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:78` |
| `New` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:90` |
| `Create` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:95` |
| `Change` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:100` |
| `Resize` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:111` |
| `AddRow` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:112` |
| `AddColumn` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:113` |
| `DeleteSelectedRowOrColumn` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:114` |
| `ReorderColumn` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:115` |
| `Refresh` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:116` |
| `Stripe` | method | `apps/powerpoint-vsto/ProcessFlowTools.cs:117` |
