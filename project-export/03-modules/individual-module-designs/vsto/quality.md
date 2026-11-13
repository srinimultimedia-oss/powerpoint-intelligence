# Read-only presentation heuristics and navigation

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Inspect bounded native content for review hints without automatic repair or accessibility certification.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/PresentationQuality.cs:1` — QualityInspectionOptions, QualityShapeSnapshot, QualitySlideSnapshot, QualityIssue, QualityInspectionReport, PresentationQuality.

## Inputs, outputs, data structures and dependencies

Inputs: presentation/QualityInspectionOptions. Outputs: report/issues/limitations and native navigation. Dependencies: native shape/text reads and pure evaluator.

## Internal execution and external integration

Inspect captures bounded slide/shape snapshots and Evaluate emits hidden-slide, off-slide, empty-text, small-text, image-alt-text and font-scan-limit issues. Unrotated top-level geometry uses.5pt tolerance; group recursion/sample limits mark Truncated. WinForms grid supports threshold/rescan/locate; nested findings locate containing parent group.

## Configuration, validation and failure behavior

Default font10pt,<=500 slides,20000 shapes,depth8,256 sampledcharacters,2000 issues. Options validate bounded alternate limits. Native COM read errors log HRESULT-only limitations.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/PresentationQualityTests.cs:11`.
- `tests/s0/WorkflowOfficeTests.cs:91`.

## Known limitations and technical debt

Masters,notes,chart-internal text,SmartArt,tablecells and rotated/group-child off-slide extents not inspected. Absence of findings is not a complete deck review; no automatic repair/contrast/brand/accessibility certification.

## Extension points

Add pure snapshot rules plus bounded native capture/tests and retain explicit incomplete scan status.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `QualityInspectionOptions` | class | `apps/powerpoint-vsto/PresentationQuality.cs:10` |
| `Validate` | method | `apps/powerpoint-vsto/PresentationQuality.cs:14` |
| `QualityShapeSnapshot` | class | `apps/powerpoint-vsto/PresentationQuality.cs:16` |
| `QualitySlideSnapshot` | class | `apps/powerpoint-vsto/PresentationQuality.cs:23` |
| `QualityIssue` | class | `apps/powerpoint-vsto/PresentationQuality.cs:30` |
| `QualityInspectionReport` | class | `apps/powerpoint-vsto/PresentationQuality.cs:35` |
| `PresentationQuality` | class | `apps/powerpoint-vsto/PresentationQuality.cs:42` |
| `Evaluate` | method | `apps/powerpoint-vsto/PresentationQuality.cs:44` |
| `Finite` | method | `apps/powerpoint-vsto/PresentationQuality.cs:74` |
| `Add` | method | `apps/powerpoint-vsto/PresentationQuality.cs:75` |
| `Inspect` | method | `apps/powerpoint-vsto/PresentationQuality.cs:80` |
| `Capture` | method | `apps/powerpoint-vsto/PresentationQuality.cs:99` |
| `ShowInspection` | method | `apps/powerpoint-vsto/PresentationQuality.cs:127` |
