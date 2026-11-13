# Native formatting plans and restoration transactions

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Apply alignment/orientation/color to bounded native text/group/table selections while retaining unrequested formatting.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/NativeFormatting.cs:1` — NativeFormattingPlan, NativeFormattingChange, NativeFormattingTransaction, NativeFormatting, Owner, Target, SelectionPlan, ColorState.

## Inputs, outputs, data structures and dependencies

Inputs: native selection, text mode or RGB. Outputs: formatting under one Undo boundary. Dependencies: PowerPoint TextFrame/TextRange/Cell/LineFormat/FillFormat; reused by context and ribbon.

## Internal execution and external integration

Collect preflights selected character ranges, group leaves and table selected cells (or whole table if none selected). Text modes map horizontal, vertical and direction settings. Color supports solid fills, outlines and per-character text color; snapshots preserve theme/tint and table borders. NativeFormattingTransaction applies changes and restores prior changes after failure.

## Configuration, validation and failure behavior

NativeFormattingPlan bounds nodes/characters/depth and limits selection500targets. RGB0–0 xFFFFFF. Unsupported/mixed nonrestorable states are rejected before mutation where possible; incomplete restoration directs Undo.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/library/LibraryBrowserTests.cs:54`.
- `tests/library/TemplateFidelity.cs:71`.
- `tests/library/TemplateFidelityTests.cs:63`.
- `tests/s0/ContextFormattingTests.cs:6`.
- `tests/s0/ContextPropertiesOfficeTests.cs:64`.
- `tests/s0/DataSourceTests.cs:27`.
- `tests/s0/NativeFormattingTests.cs:11`.
- `tests/s0/TextShapeOfficeTests.cs:31`.
- `tests/s0/WorkflowOfficeTests.cs:88`.
- `tests/library/piidx/PiidxEdgeTests.cs:18`.
- `tests/library/piidx/PiidxSourceTests.cs:24`.

## Known limitations and technical debt

SmartArt, arbitrary embedded objects, gradient/pattern/texture fills and advanced rich-text document serialization are not implemented as generic conversions.

## Extension points

Add properties as validated snapshot/apply/restore operations with group/table/selected-character tests.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `NativeFormattingPlan` | class | `apps/powerpoint-vsto/NativeFormatting.cs:12` |
| `ColorKind` | method | `apps/powerpoint-vsto/NativeFormatting.cs:15` |
| `TextMode` | method | `apps/powerpoint-vsto/NativeFormatting.cs:16` |
| `ColorValue` | method | `apps/powerpoint-vsto/NativeFormatting.cs:17` |
| `Rgb` | method | `apps/powerpoint-vsto/NativeFormatting.cs:18` |
| `Bounds` | method | `apps/powerpoint-vsto/NativeFormatting.cs:19` |
| `HorizontalAlignment` | method | `apps/powerpoint-vsto/NativeFormatting.cs:20` |
| `VerticalAlignment` | method | `apps/powerpoint-vsto/NativeFormatting.cs:21` |
| `Paragraph` | method | `apps/powerpoint-vsto/NativeFormatting.cs:22` |
| `Anchor` | method | `apps/powerpoint-vsto/NativeFormatting.cs:23` |
| `Orientation` | method | `apps/powerpoint-vsto/NativeFormatting.cs:24` |
| `NativeFormattingChange` | class | `apps/powerpoint-vsto/NativeFormatting.cs:26` |
| `NativeFormattingTransaction` | class | `apps/powerpoint-vsto/NativeFormatting.cs:31` |
| `Run` | method | `apps/powerpoint-vsto/NativeFormatting.cs:33` |
| `NativeFormatting` | class | `apps/powerpoint-vsto/NativeFormatting.cs:40` |
| `Owner` | class | `apps/powerpoint-vsto/NativeFormatting.cs:42` |
| `Target` | class | `apps/powerpoint-vsto/NativeFormatting.cs:43` |
| `SelectionPlan` | class | `apps/powerpoint-vsto/NativeFormatting.cs:44` |
| `ColorState` | class | `apps/powerpoint-vsto/NativeFormatting.cs:45` |
| `Restore` | method | `apps/powerpoint-vsto/NativeFormatting.cs:49` |
| `CanApply` | method | `apps/powerpoint-vsto/NativeFormatting.cs:51` |
| `SelectedShapes` | method | `apps/powerpoint-vsto/NativeFormatting.cs:53` |
| `Selection` | method | `apps/powerpoint-vsto/NativeFormatting.cs:54` |
| `Geometry` | method | `apps/powerpoint-vsto/NativeFormatting.cs:56` |
| `Add` | method | `apps/powerpoint-vsto/NativeFormatting.cs:58` |
| `Collect` | method | `apps/powerpoint-vsto/NativeFormatting.cs:66` |
| `Color` | method | `apps/powerpoint-vsto/NativeFormatting.cs:72` |
| `ApplyColor` | method | `apps/powerpoint-vsto/NativeFormatting.cs:76` |
| `Outline` | method | `apps/powerpoint-vsto/NativeFormatting.cs:86` |
| `Text` | method | `apps/powerpoint-vsto/NativeFormatting.cs:88` |
