# Local symbol, decimal, picture-copy and chart workbook tools

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Provide bounded local selection tools with explicit user intent and original-object retention.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/ProductivityTools.cs:1` — ProductivitySymbol, DecimalCharacterChange, ProductivityTools, TextEdit, WorkbookItem.

## Inputs, outputs, data structures and dependencies

Inputs: symbol/query, native selection/confirmation. Outputs: editable text, picture copy or open embedded chart workbook. Dependencies: PowerPoint COM, System.Drawing and dialogs; no email/cloud asset service.

## Internal execution and external integration

Symbol picker searches finite Unicode catalog and inserts/replaces selected text or adds textbox. Decimal tool plans digit-adjacent dots/commas across selected text/group leaves, confirms, changes individual characters and restores on failure. RasterCopy accepts one unrotated object/group, exports/validates PNG, embeds picture at same bounds and retains original. Workbook list navigates top-level chart objects and opens each separate embedded Edit Data workbook. Routed manager/style/save/review aliases are handled elsewhere.

## Configuration, validation and failure behavior

Query<=100 chars; decimal<=20000 chars/100shape visits/depth8; image<=4096 px/side and64MiB; chart list<=500 slides/2000shapes. Temporary raster folder is removed when possible.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/DemoContractTests.cs:50`.
- `tests/s0/ProductivityToolOfficeTests.cs:14`.
- `tests/s0/ProductivityToolTests.cs:10`.

## Known limitations and technical debt

Table cells/chart labels excluded from decimal substitution; digit-separated dates/thousands separators also change after confirmation. Raster is intentionally a picture copy. No shared slide workbook exists. Cleanup/rollback failures are visible rather than swallowed as successful edits.

## Extension points

Add tools with explicit budgets/selection semantics, retention/confirmation policy and native/source regression tests.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `ProductivitySymbol` | class | `apps/powerpoint-vsto/ProductivityTools.cs:11` |
| `ToString` | method | `apps/powerpoint-vsto/ProductivityTools.cs:15` |
| `DecimalCharacterChange` | class | `apps/powerpoint-vsto/ProductivityTools.cs:17` |
| `ProductivityTools` | class | `apps/powerpoint-vsto/ProductivityTools.cs:21` |
| `SearchSymbols` | method | `apps/powerpoint-vsto/ProductivityTools.cs:27` |
| `DecimalChanges` | method | `apps/powerpoint-vsto/ProductivityTools.cs:29` |
| `Run` | method | `apps/powerpoint-vsto/ProductivityTools.cs:34` |
| `ActiveSlide` | method | `apps/powerpoint-vsto/ProductivityTools.cs:39` |
| `InsertSymbol` | method | `apps/powerpoint-vsto/ProductivityTools.cs:44` |
| `ShowSymbols` | method | `apps/powerpoint-vsto/ProductivityTools.cs:50` |
| `TextEdit` | class | `apps/powerpoint-vsto/ProductivityTools.cs:59` |
| `CollectText` | method | `apps/powerpoint-vsto/ProductivityTools.cs:60` |
| `SwitchDecimals` | method | `apps/powerpoint-vsto/ProductivityTools.cs:66` |
| `RasterCopy` | method | `apps/powerpoint-vsto/ProductivityTools.cs:79` |
| `WorkbookItem` | class | `apps/powerpoint-vsto/ProductivityTools.cs:90` |
| `ToString` | method | `apps/powerpoint-vsto/ProductivityTools.cs:93` |
| `ShowWorkbooks` | method | `apps/powerpoint-vsto/ProductivityTools.cs:95` |
