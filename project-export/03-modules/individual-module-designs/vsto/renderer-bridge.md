# Native scene-renderer ribbon adapter

**VERIFIED — source implementation; native runtime acceptance not re-executed in this documentation review.**

## Purpose/status

Adapt validated canonical data into a deterministic clustered-column scene and insert native editable grouped primitives.

## Sources/classes/interfaces

`apps/powerpoint-vsto/RendererRibbonCommand.cs:11` defines the static adapter. It connects the source-proven core rendering contracts documented separately.

## Inputs/outputs/dependencies

Inputs: embedded fixture pair for explicit sample/helper or linked ChartDocument. Outputs: ValidatedProjection and NativeRenderOutcome. Dependencies: ChartCodec, ClusteredProjection, SceneCodec, NativeSceneRenderer, PowerPointSceneSurface.

## Execution/external integration

Prepare() validates the embedded canonical/scene pair before assigning new identity, remaps node/diagnostic IDs and recomputes render hash. The active ribbon path calls Prepare(model): validates the supplied linked v2 model, computes signed value bounds including zero, skips missing values, emits rectangle bars and diagnostics, then validates model/scene consistency. Insert binds NativeSceneRenderer to a PowerPointSceneSurface for the existing active deck/slide; no new presentation is created.

## Configuration/validation

Finite numeric range required. Scene/model IDs/revision/units/hash must agree. Current insertion requires an active Normal-view slide. Metadata is owned by core rendering surface, not the S0 chart CustomerData contract.

## Failure handling, security and ownership

Source/model validation precedes COM. Invalid resource pair, missing active slide or infinite range throws. NativeSceneRenderer owns rollback of newly created primitives. There is no executable AI payload or arbitrary external COM command deserialization. The caller supplies a native undo/mutation boundary. Sample resources are embedded and do not require a writable filesystem sample file.

## Tests

`tests/s3/RendererRibbonTests.cs:1` and `tests/s3/NativeRendererTests.cs:1` contain renderer fixture/projection/native adapter tests; source references do not prove this machine executed Office.

## Limitations/debt

Active linked renderer emits AXES_NOT_RENDERED and optionally LEGEND_NOT_RENDERED. The native object is a shape group, not a PowerPoint Chart. Embedded sample Prepare() remains an internal helper/test route and should not be confused with the user-linked ribbon path. Not every canonical chart type is projected here.

## Extension points

Add new chart-family projections only with canonical-scene equivalence, primitive render/hash and native edit/save-reopen acceptance.

## Important methods

| Method | Source |
|---|---|
| Prepare() validates original sample then remaps identity | `apps/powerpoint-vsto/RendererRibbonCommand.cs:21` |
| Prepare(model) creates linked rectangle projection | `apps/powerpoint-vsto/RendererRibbonCommand.cs:44` |
| Insert(application,model) binds existing deck and slide | `apps/powerpoint-vsto/RendererRibbonCommand.cs:60` |
