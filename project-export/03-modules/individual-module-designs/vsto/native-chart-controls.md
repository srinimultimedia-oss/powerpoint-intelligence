# Native axes, series and annotations

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

Edit native axes/series and add editable reference/difference annotations without flattening charts.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/NativeChartControls.cs:1` — NativeAxisSettings, NativeChartControls, Snapshot.
- `apps/powerpoint-vsto/NativeSeriesControls.cs:1` — NativeSeriesControls, NativePdfExport.
- `apps/powerpoint-vsto/ChartAnnotations.cs:1` — ChartAnnotationSpec, ChartAnnotationPlan, ChartAnnotations, Context.

## Inputs, outputs, data structures and dependencies

Inputs: selected charts, NativeAxisSettings, series index/type/group, annotation value/indices or PDF destination. Outputs: native chart settings, editable annotation shapes or PDF. Dependencies: PowerPoint Chart/Axis/SeriesCollection, geometry plans and rollback patterns.

## Internal execution and external integration

Axis settings snapshot primary axis, validate finite bounds/number format and positive log data, reset auto bounds then apply and restore on error. MatchScales preflights2–30 charts with matching linear/log type and preserves individual reversal/format. Series controls allow existing unstacked2D column/line/line-markers/area with primary/secondary grouping, but reject S0 metadata. Annotations map reference value or selected point differences through native plot/axis geometry, draw owned groups and refresh explicitly. NativePdfExport exports all slides with ExportAsFixedFormat.

## Configuration, validation and failure behavior

NumberFormat1–100 chars; series<=100; scale matching2–30 charts. Log preflight requires all present primary values positive. Native chart capabilities determine axis availability. StartNewUndoEntry precedes mutation.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/library/GalleryTests.cs:28`.
- `tests/s0/ChartAnnotationTests.cs:35`.
- `tests/s0/ContextPanelModelTests.cs:7`.
- `tests/s0/ContextPropertiesOfficeTests.cs:54`.
- `tests/s0/ContextPropertiesTests.cs:15`.
- `tests/s0/DemoContractTests.cs:22`.
- `tests/s0/DemoScenario.cs:170`.
- `tests/s0/NativeChartControlTests.cs:13`.
- `tests/s0/WorkflowOfficeTests.cs:72`.

## Known limitations and technical debt

Unsupported types/axes, date-category difference geometry and out-of-range annotations are rejected. Annotations are grouped shapes, not chart-internal adornments, and require refresh. Complete automatic label collision/layout optimization is not implemented. Rollback errors aggregate and request Undo.

## Extension points

Extend explicit capability predicates and pure plans before widening native adapters; add geometry, rollback, secondary-axis and save/reopen scenarios.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `NativeAxisSettings` | class | `apps/powerpoint-vsto/NativeChartControls.cs:10` |
| `Validate` | method | `apps/powerpoint-vsto/NativeChartControls.cs:15` |
| `Finite` | method | `apps/powerpoint-vsto/NativeChartControls.cs:22` |
| `ValidateLogValues` | method | `apps/powerpoint-vsto/NativeChartControls.cs:23` |
| `NativeChartControls` | class | `apps/powerpoint-vsto/NativeChartControls.cs:29` |
| `Selected` | method | `apps/powerpoint-vsto/NativeChartControls.cs:31` |
| `Axis` | method | `apps/powerpoint-vsto/NativeChartControls.cs:33` |
| `Snapshot` | class | `apps/powerpoint-vsto/NativeChartControls.cs:35` |
| `Restore` | method | `apps/powerpoint-vsto/NativeChartControls.cs:39` |
| `Apply` | method | `apps/powerpoint-vsto/NativeChartControls.cs:41` |
| `SetAxis` | method | `apps/powerpoint-vsto/NativeChartControls.cs:49` |
| `ShowAxis` | method | `apps/powerpoint-vsto/NativeChartControls.cs:55` |
| `MatchScales` | method | `apps/powerpoint-vsto/NativeChartControls.cs:69` |
| `NativeSeriesControls` | class | `apps/powerpoint-vsto/NativeSeriesControls.cs:9` |
| `Validate` | method | `apps/powerpoint-vsto/NativeSeriesControls.cs:12` |
| `Apply` | method | `apps/powerpoint-vsto/NativeSeriesControls.cs:13` |
| `Show` | method | `apps/powerpoint-vsto/NativeSeriesControls.cs:20` |
| `NativePdfExport` | class | `apps/powerpoint-vsto/NativeSeriesControls.cs:31` |
| `Export` | method | `apps/powerpoint-vsto/NativeSeriesControls.cs:33` |
| `Show` | method | `apps/powerpoint-vsto/NativeSeriesControls.cs:39` |
| `ChartAnnotationSpec` | class | `apps/powerpoint-vsto/ChartAnnotations.cs:11` |
| `Validate` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:14` |
| `Tags` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:15` |
| `Read` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:16` |
| `ChartAnnotationPlan` | class | `apps/powerpoint-vsto/ChartAnnotations.cs:18` |
| `Number` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:21` |
| `Finite` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:22` |
| `Bounds` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:23` |
| `Position` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:24` |
| `Reference` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:25` |
| `Difference` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:26` |
| `ChartAnnotations` | class | `apps/powerpoint-vsto/ChartAnnotations.cs:28` |
| `Context` | class | `apps/powerpoint-vsto/ChartAnnotations.cs:31` |
| `Selected` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:32` |
| `Inspect` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:33` |
| `Values` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:38` |
| `Plan` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:39` |
| `Canvas` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:40` |
| `Draw` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:41` |
| `Reference` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:45` |
| `Difference` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:46` |
| `Insert` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:47` |
| `RefreshSelected` | method | `apps/powerpoint-vsto/ChartAnnotations.cs:51` |
