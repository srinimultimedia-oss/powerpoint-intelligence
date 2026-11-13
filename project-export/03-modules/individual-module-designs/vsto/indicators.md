# Harvey-ball and checkbox editable vectors

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Render editable grouped indicators with typed canonical state, conversion, colors, rotation and focused keyboard shortcuts.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/IndicatorElements.cs:1` — IndicatorSpec, IndicatorCodec, IndicatorElements.

## Inputs, outputs, data structures and dependencies

Inputs: IndicatorSpec and native group selection. Output: tagged editable group and XML. Dependencies: native shapes/freeform/grouping, XmlPartUpdates attachment lookup and ElementCommands keyboard form.

## Internal execution and external integration

Render creates owned native background/outline/freeform-sector or checkbox strokes/text, groups and attaches IndicatorCodec XML. Configure renders replacement at supported original bounds and then removes old group. Step clamps Harvey fill or cycles three checkbox states. Numeric shortcuts map fill/state; focused keyboard controls re-find original slide/shape; double-click intercepts only owned indicators.

## Configuration, validation and failure behavior

Segments1–100,fill0–segments,checkbox0/1/2,size6–720 pt,rotation[0,360),customsymbol<=8 printablechars,RGB24bit; groups expect2–6 ownedchildren. Bounds must fit slide.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/ContextPropertiesOfficeTests.cs:68`.
- `tests/s0/DemoScenario.cs:145`.
- `tests/s0/IndicatorElementTests.cs:6`.
- `tests/s0/IndicatorOfficeTests.cs:14`.

## Known limitations and technical debt

Canonical updates replace managed group, not foreign child content. External data-linked indicators are absent; SetValueSelected is a direct local helper. Unowned children/ambiguous XML fail preserved; incomplete replacement cleanup retains objects and requests Undo.

## Extension points

Add indicator kinds through versioned state/owned vectors/keyboard semantics and conversion/native-persistence tests.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `IndicatorSpec` | class | `apps/powerpoint-vsto/IndicatorElements.cs:11` |
| `Copy` | method | `apps/powerpoint-vsto/IndicatorElements.cs:13` |
| `Validate` | method | `apps/powerpoint-vsto/IndicatorElements.cs:14` |
| `Step` | method | `apps/powerpoint-vsto/IndicatorElements.cs:16` |
| `Shortcut` | method | `apps/powerpoint-vsto/IndicatorElements.cs:17` |
| `IndicatorCodec` | class | `apps/powerpoint-vsto/IndicatorElements.cs:19` |
| `Encode` | method | `apps/powerpoint-vsto/IndicatorElements.cs:21` |
| `Decode` | method | `apps/powerpoint-vsto/IndicatorElements.cs:22` |
| `IndicatorElements` | class | `apps/powerpoint-vsto/IndicatorElements.cs:25` |
| `IsIndicator` | method | `apps/powerpoint-vsto/IndicatorElements.cs:27` |
| `Read` | method | `apps/powerpoint-vsto/IndicatorElements.cs:29` |
| `Insert` | method | `apps/powerpoint-vsto/IndicatorElements.cs:30` |
| `Configure` | method | `apps/powerpoint-vsto/IndicatorElements.cs:32` |
| `Step` | method | `apps/powerpoint-vsto/IndicatorElements.cs:33` |
| `SetShortcut` | method | `apps/powerpoint-vsto/IndicatorElements.cs:34` |
| `CycleSelected` | method | `apps/powerpoint-vsto/IndicatorElements.cs:35` |
| `ConvertSelected` | method | `apps/powerpoint-vsto/IndicatorElements.cs:36` |
| `SetValueSelected` | method | `apps/powerpoint-vsto/IndicatorElements.cs:37` |
| `Edit` | method | `apps/powerpoint-vsto/IndicatorElements.cs:38` |
