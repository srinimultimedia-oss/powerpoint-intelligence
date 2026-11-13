# Selection-aware floating property box

<!-- generated-toc -->
## Contents

- [Purpose and implementation status](#purpose-and-implementation-status)
- [Source locations and internal components](#source-locations-and-internal-components)
- [Inputs, outputs, data structures and dependencies](#inputs-outputs-data-structures-and-dependencies)
- [Internal execution and external integration](#internal-execution-and-external-integration)
- [Configuration, validation and failure behavior](#configuration-validation-and-failure-behavior)
- [Security considerations and cleanup](#security-considerations-and-cleanup)
- [Test references](#test-references)
- [Known limitations and technical debt](#known-limitations-and-technical-debt)
- [Extension points](#extension-points)
- [Classes, interfaces and method navigation](#classes-interfaces-and-method-navigation)
<!-- end-generated-toc -->

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Inspect live selection and expose applicable native properties/actions without stealing focus or editing stale targets.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/ContextPanelModel.cs:1` — ContextKind, ContextDescriptor, ContextSnapshot, MixedValueKind, MixedValue, ContextPanelModel.
- `apps/powerpoint-vsto/ContextSelectionBridge.cs:1` — ContextNativeState, ContextSelectionBridge, Rect, Targets.
- `apps/powerpoint-vsto/ContextPropertiesController.cs:1` — ContextPropertiesController, Owner.
- `apps/powerpoint-vsto/ContextPropertiesView.cs:1` — PropertySwatch, ContextPropertiesView.
- `apps/powerpoint-vsto/ContextFormatting.cs:1` — FontSelectionState, ContextFontPlan, ContextFormatting, Collector.
- `apps/powerpoint-vsto/ContextChartFormatting.cs:1` — ContextChartFontPlan, ContextChartFormatting.
- `apps/powerpoint-vsto/ContextActionRouter.cs:1` — ContextActionRouter.

## Inputs, outputs, data structures and dependencies

Inputs: selection, owner/foreground metrics and user edits. Outputs: ContextSnapshot/ContextNativeState, floating WinForms controls and native changes. Dependencies: user32 geometry/foreground calls, balanced IUnknown pointer release, NativeFormatting and managed adapters.

## Internal execution and external integration

A300ms UI timer plus selection dirtiness checks availability/foreground. Selection identity includes presentation IUnknown/window/slide/shapes/text or table cells. Capture aggregates capabilities/mixed values, rechecks identity and places an owned modeless form within monitor/window bounds. Every property edit revalidates the key, including after ColorDialog. Managed indicator/connector color changes update their canonical state; chart fonts/fill use chart-specific adapters. Only requested font properties are mutated; volatile XML observation is produced on UI and read from automation without COM.

## Configuration, validation and failure behavior

Inspection<=500 native parts/selected objects, group depth16,<=2000 text characters and<=400 tablecells. Read-only/slideshow/no-supported-selection hides or disables controls. Mixed property values stay mixed. Exceptions report type/HRESULT/stage rather than document values.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/ContextChartFormattingTests.cs:4`.
- `tests/s0/ContextFormattingOfficeTests.cs:8`.
- `tests/s0/ContextFormattingTests.cs:6`.
- `tests/s0/ContextPanelModelTests.cs:13`.
- `tests/s0/ContextPropertiesOfficeTests.cs:24`.
- `tests/s0/ContextPropertiesTests.cs:15`.

## Known limitations and technical debt

Per-monitor DPI, native color picker, every chart family and all focus interactions need desktop manual coverage. Overview fallback may classify while withholding edits. Inspection bounds can make larger selections noneditable; this is not complete reference-product contextual-toolbar parity.

## Extension points

Add descriptors/capabilities/collection/action routing consistently and test stale selection, read-only, mixed state, native formatting and focus preservation.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `ContextKind` | enum | `apps/powerpoint-vsto/ContextPanelModel.cs:10` |
| `ContextDescriptor` | class | `apps/powerpoint-vsto/ContextPanelModel.cs:11` |
| `ContextSnapshot` | class | `apps/powerpoint-vsto/ContextPanelModel.cs:19` |
| `MixedValueKind` | enum | `apps/powerpoint-vsto/ContextPanelModel.cs:31` |
| `MixedValue` | class | `apps/powerpoint-vsto/ContextPanelModel.cs:32` |
| `From` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:39` |
| `ContextPanelModel` | class | `apps/powerpoint-vsto/ContextPanelModel.cs:44` |
| `Managed` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:49` |
| `Classify` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:50` |
| `Actions` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:57` |
| `Build` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:62` |
| `SelectionKey` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:74` |
| `IsCurrent` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:79` |
| `CanExecute` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:80` |
| `RectangleBounds` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:81` |
| `Clamp` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:85` |
| `Area` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:86` |
| `Place` | method | `apps/powerpoint-vsto/ContextPanelModel.cs:87` |
| `ContextNativeState` | class | `apps/powerpoint-vsto/ContextSelectionBridge.cs:12` |
| `ContextSelectionBridge` | class | `apps/powerpoint-vsto/ContextSelectionBridge.cs:22` |
| `Rect` | struct | `apps/powerpoint-vsto/ContextSelectionBridge.cs:25` |
| `Shapes` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:29` |
| `ObjectKey` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:30` |
| `Key` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:31` |
| `Ensure` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:39` |
| `Targets` | class | `apps/powerpoint-vsto/ContextSelectionBridge.cs:40` |
| `TargetsOf` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:41` |
| `Common` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:51` |
| `Common` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:52` |
| `Frame` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:53` |
| `Paragraph` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:54` |
| `Vertical` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:55` |
| `Direction` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:56` |
| `Capture` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:57` |
| `ApplyColor` | method | `apps/powerpoint-vsto/ContextSelectionBridge.cs:86` |
| `ContextPropertiesController` | class | `apps/powerpoint-vsto/ContextPropertiesController.cs:9` |
| `Owner` | class | `apps/powerpoint-vsto/ContextPropertiesController.cs:27` |
| `Request` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:30` |
| `Hide` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:31` |
| `ShowCurrent` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:32` |
| `Refresh` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:33` |
| `IsInspectionFailure` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:57` |
| `Failure` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:58` |
| `Apply` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:59` |
| `PublishObservation` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:84` |
| `AutomationState` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:88` |
| `Dispose` | method | `apps/powerpoint-vsto/ContextPropertiesController.cs:89` |
| `PropertySwatch` | class | `apps/powerpoint-vsto/ContextPropertiesView.cs:8` |
| `OnPaint` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:11` |
| `ContextPropertiesView` | class | `apps/powerpoint-vsto/ContextPropertiesView.cs:13` |
| `Choices` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:42` |
| `SizeCommand` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:43` |
| `Send` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:44` |
| `Dismiss` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:45` |
| `Select` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:46` |
| `Render` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:47` |
| `Preview` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:54` |
| `Descendants` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:64` |
| `Error` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:65` |
| `Caption` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:66` |
| `Dispose` | method | `apps/powerpoint-vsto/ContextPropertiesView.cs:67` |
| `FontSelectionState` | class | `apps/powerpoint-vsto/ContextFormatting.cs:8` |
| `ContextFontPlan` | class | `apps/powerpoint-vsto/ContextFormatting.cs:9` |
| `Validate` | method | `apps/powerpoint-vsto/ContextFormatting.cs:10` |
| `Merge` | method | `apps/powerpoint-vsto/ContextFormatting.cs:11` |
| `ContextFormatting` | class | `apps/powerpoint-vsto/ContextFormatting.cs:13` |
| `Collector` | class | `apps/powerpoint-vsto/ContextFormatting.cs:14` |
| `Read` | method | `apps/powerpoint-vsto/ContextFormatting.cs:20` |
| `ApplyFont` | method | `apps/powerpoint-vsto/ContextFormatting.cs:21` |
| `ContextChartFontPlan` | class | `apps/powerpoint-vsto/ContextChartFormatting.cs:8` |
| `Family` | method | `apps/powerpoint-vsto/ContextChartFormatting.cs:9` |
| `Size` | method | `apps/powerpoint-vsto/ContextChartFormatting.cs:10` |
| `Flag` | method | `apps/powerpoint-vsto/ContextChartFormatting.cs:11` |
| `Color` | method | `apps/powerpoint-vsto/ContextChartFormatting.cs:12` |
| `ContextChartFormatting` | class | `apps/powerpoint-vsto/ContextChartFormatting.cs:14` |
| `Read` | method | `apps/powerpoint-vsto/ContextChartFormatting.cs:16` |
| `ApplyFont` | method | `apps/powerpoint-vsto/ContextChartFormatting.cs:17` |
| `ReadTextColor` | method | `apps/powerpoint-vsto/ContextChartFormatting.cs:19` |
| `ApplyColor` | method | `apps/powerpoint-vsto/ContextChartFormatting.cs:20` |
| `ContextActionRouter` | class | `apps/powerpoint-vsto/ContextActionRouter.cs:7` |
| `Run` | method | `apps/powerpoint-vsto/ContextActionRouter.cs:9` |
