# Text box, pentagon, chevron and rounded native elements

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
- [Length and edge-lock contract examples](#length-and-edge-lock-contract-examples)
<!-- end-generated-toc -->

**VERIFIED — source implementation. Current runtime acceptance is not established by this design review.**

## Purpose and implementation status

Provide editable shapes with canonical geometry/edge locks, fit, convert, adjacent duplication and contextual settings.

Implemented local source is distinguished from supported subset/absent external behavior below. See current test evidence in the export testing chapter; do not treat historical logs or method existence as a current runtime pass.

## Source locations and internal components

- `apps/powerpoint-vsto/TextShapeElements.cs:1` — TextEdgeLocks, TextShapeSpec, TextShapePlan, TextShapeCodec, TextShapeElements.
- `apps/powerpoint-vsto/ElementCommands.cs:1` — ElementCommands.

## Inputs, outputs, data structures and dependencies

Inputs: kind, TextShapeSpec, edge locks/adjustment/autofit and selection. Outputs: native editable shape, PPI_TEXT_ELEMENT and versioned XML. Dependencies: TextShapePlan/Codec, native metrics, XmlPartUpdates and lifecycle timers.

## Internal execution and external integration

Insert creates a centered200x70pt native AutoShape, SegoeUI20, center/middle alignment, manual sizing/margins and canonical shape XML/tag. Configure validates kind/adjustment and measures text against unlocked edges and slide bounds, then replaces XML and geometry with rollback. FitSelected precomputes1–100 managed shape plans. Conversion/duplication preserve supported native formatting and independent metadata. ElementCommands gates Normal-view/current ownership; automatic fit runs via add-in timer.

## Configuration, validation and failure behavior

Edge flags Left/Right/Top/Bottom; adjustment.05–.5. Geometry must be supported unrotated/unflipped/native orientation. All-locked fit requires content to fit existing bounds. Temporary preflight duplicates never attach canonical metadata.

## Security considerations and cleanup

Local native calls run in the current Office user process and require supported selection/ownership; a mutex/selection key is not authentication or RBAC. File/metadata content must remain bounded and error text must not expose data or credentials. Explicit owned-resource cleanup and error/rollback paths are part of the cited implementation, not a guarantee that a failed COM server can always be restored.

## Test references

Lexical module-class matches in test source (not instrumentation coverage; no test is claimed executed here):

- `tests/s0/ContextPropertiesOfficeTests.cs:66`.
- `tests/s0/DemoContractTests.cs:49`.
- `tests/s0/DemoScenario.cs:142`.
- `tests/s0/ProcessFlowOfficeTests.cs:20`.
- `tests/s0/TextShapeElementTests.cs:27`.
- `tests/s0/TextShapeOfficeTests.cs:21`.

## Known limitations and technical debt

Rotated/flipped/manually changed-kind objects and copied grid cells can fail closed. Text layout uses native metrics, not independent typography. Missing/ambiguous XML and unfittable geometry are preserved/rejected; incomplete rollback surfaces Undo.

## Extension points

Add shape kinds through native mapping, spec/codec, fit/convert geometry, commands/icons and native persistence tests.

## Classes, interfaces and method navigation

Method names below are source declarations; private/internal members are not public API promises. The source hash/review manifest gives exact reviewed bytes.

| Declaration | Kind | Source |
|---|---|---|
| `TextEdgeLocks` | enum | `apps/powerpoint-vsto/TextShapeElements.cs:15` |
| `TextShapeSpec` | class | `apps/powerpoint-vsto/TextShapeElements.cs:16` |
| `Copy` | method | `apps/powerpoint-vsto/TextShapeElements.cs:22` |
| `TextShapePlan` | class | `apps/powerpoint-vsto/TextShapeElements.cs:24` |
| `Kind` | method | `apps/powerpoint-vsto/TextShapeElements.cs:26` |
| `Finite` | method | `apps/powerpoint-vsto/TextShapeElements.cs:32` |
| `Validate` | method | `apps/powerpoint-vsto/TextShapeElements.cs:33` |
| `ParseLength` | method | `apps/powerpoint-vsto/TextShapeElements.cs:34` |
| `Fit` | method | `apps/powerpoint-vsto/TextShapeElements.cs:45` |
| `TextShapeCodec` | class | `apps/powerpoint-vsto/TextShapeElements.cs:55` |
| `Encode` | method | `apps/powerpoint-vsto/TextShapeElements.cs:58` |
| `Decode` | method | `apps/powerpoint-vsto/TextShapeElements.cs:59` |
| `TextShapeElements` | class | `apps/powerpoint-vsto/TextShapeElements.cs:68` |
| `IsManaged` | method | `apps/powerpoint-vsto/TextShapeElements.cs:71` |
| `Active` | method | `apps/powerpoint-vsto/TextShapeElements.cs:72` |
| `Selected` | method | `apps/powerpoint-vsto/TextShapeElements.cs:73` |
| `Geometry` | method | `apps/powerpoint-vsto/TextShapeElements.cs:74` |
| `Native` | method | `apps/powerpoint-vsto/TextShapeElements.cs:75` |
| `Part` | method | `apps/powerpoint-vsto/TextShapeElements.cs:76` |
| `Read` | method | `apps/powerpoint-vsto/TextShapeElements.cs:77` |
| `Bounds` | method | `apps/powerpoint-vsto/TextShapeElements.cs:78` |
| `Bounds` | method | `apps/powerpoint-vsto/TextShapeElements.cs:79` |
| `Adjust` | method | `apps/powerpoint-vsto/TextShapeElements.cs:80` |
| `Measure` | method | `apps/powerpoint-vsto/TextShapeElements.cs:81` |
| `Plan` | method | `apps/powerpoint-vsto/TextShapeElements.cs:89` |
| `InitializeNewShape` | method | `apps/powerpoint-vsto/TextShapeElements.cs:90` |
| `Insert` | method | `apps/powerpoint-vsto/TextShapeElements.cs:103` |
| `PreflightGridGeometry` | method | `apps/powerpoint-vsto/TextShapeElements.cs:107` |
| `Configure` | method | `apps/powerpoint-vsto/TextShapeElements.cs:117` |
| `FitSelected` | method | `apps/powerpoint-vsto/TextShapeElements.cs:121` |
| `DuplicateSelected` | method | `apps/powerpoint-vsto/TextShapeElements.cs:126` |
| `ConvertSelected` | method | `apps/powerpoint-vsto/TextShapeElements.cs:134` |
| `FitSnapshot` | method | `apps/powerpoint-vsto/TextShapeElements.cs:139` |
| `AutoReflow` | method | `apps/powerpoint-vsto/TextShapeElements.cs:141` |
| `DialogAction` | method | `apps/powerpoint-vsto/TextShapeElements.cs:145` |
| `Edit` | method | `apps/powerpoint-vsto/TextShapeElements.cs:146` |
| `ElementCommands` | class | `apps/powerpoint-vsto/ElementCommands.cs:11` |
| `Enabled` | method | `apps/powerpoint-vsto/ElementCommands.cs:15` |
| `Run` | method | `apps/powerpoint-vsto/ElementCommands.cs:20` |
| `Stripe` | method | `apps/powerpoint-vsto/ElementCommands.cs:48` |
| `Resize` | method | `apps/powerpoint-vsto/ElementCommands.cs:50` |
| `IndicatorKeys` | method | `apps/powerpoint-vsto/ElementCommands.cs:52` |

## Length and edge-lock contract examples

The pure parser supports points, mm, cm, inches, explicit relative slide units/percentages and fractions of slide dimensions. Tests establish `25.4mm` and `1 IN` as72pt, `20%w` as192pt on960ptwidth and `1/2 sh` as270pt on540ptheight (`tests/s0/TextShapeElementTests.cs:13`). Unitless input is points. Ambiguous pixel/scientific/comma-decimal formats, zero denominator and fractions of physical units are rejected. Unlocked edges grow or shrink according to measured native text; fully locked bounds stay unchanged only when content fits. This typed parser never evaluates an expression or script string.
