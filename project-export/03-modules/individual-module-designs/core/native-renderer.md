# Canonical projection and owned native shape renderer

<!-- generated-toc -->
## Contents

- [1. Purpose](#1-purpose)
- [2. Implementation status](#2-implementation-status)
- [3. Source locations](#3-source-locations)
- [4. Internal components](#4-internal-components)
- [5. Classes and interfaces](#5-classes-and-interfaces)
- [6. Important methods and functions](#6-important-methods-and-functions)
- [7. Inputs and outputs](#7-inputs-and-outputs)
- [8. Dependencies](#8-dependencies)
- [9. Workflow](#9-workflow)
- [10. Integration points](#10-integration-points)
- [11. Exceptions and recovery](#11-exceptions-and-recovery)
- [12. Configuration dependencies](#12-configuration-dependencies)
- [13. Security considerations](#13-security-considerations)
- [14. Unit and integration test definitions](#14-unit-and-integration-test-definitions)
- [15. Known limitations](#15-known-limitations)
- [16. Technical debt](#16-technical-debt)
- [17. Extension points](#17-extension-points)
<!-- end-generated-toc -->

**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.

## 1. Purpose

Validates canonical clustered data against scene then creates editable native rectangles grouped with embedded canonical/scene/bindings.

## 2. Implementation status

Compiled bounded feasibility path; native runtime evidence remains separate from unit/source proof.

## 3. Source locations

- `src/Chart.Renderers/NativeRenderingCore.cs:1`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:1`

## 4. Internal components

ChartCodec/SceneCodec, XmlReader/Writer, PowerPoint and Office COM; owner STA/thread. Before Begin validate ID/revision/layout, ordered numeric bars/paint and omission diagnostics; cap2000 nodes/minimum 2 pt plot. Create two invisible anchors; omit zero-area native shape but retain binding. Group only operation-owned children, persist verified CustomerData XML, complete. Any mutation exception rolls back owned objects and reports rollback status.

## 5. Classes and interfaces

- `src/Chart.Renderers/NativeRenderingCore.cs:13` — `public sealed class ValidatedProjection`
- `src/Chart.Renderers/NativeRenderingCore.cs:18` — `public static class ClusteredProjection`
- `src/Chart.Renderers/NativeRenderingCore.cs:59` — `public sealed class NativeBinding`
- `src/Chart.Renderers/NativeRenderingCore.cs:65` — `public sealed class RenderedChartRecord`
- `src/Chart.Renderers/NativeRenderingCore.cs:71` — `public static class RenderedChartMetadata`
- `src/Chart.Renderers/NativeRenderingCore.cs:136` — `public interface IEditableSceneSurface`
- `src/Chart.Renderers/NativeRenderingCore.cs:147` — `public sealed class NativeRenderOutcome`
- `src/Chart.Renderers/NativeRenderingCore.cs:152` — `public sealed class NativeRenderException : Exception`
- `src/Chart.Renderers/NativeRenderingCore.cs:158` — `public sealed class NativeSceneRenderer`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:14` — `public sealed class PowerPointSceneSurface : IEditableSceneSurface`

## 6. Important methods and functions

Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.

- `src/Chart.Renderers/NativeRenderingCore.cs:23` — `public static ValidatedProjection Validate(string canonicalJson, string sceneJson)`
- `src/Chart.Renderers/NativeRenderingCore.cs:56` — `public static bool HasArea(SceneNode node)`
- `src/Chart.Renderers/NativeRenderingCore.cs:75` — `public static string NodeToken(string id)`
- `src/Chart.Renderers/NativeRenderingCore.cs:78` — `public static ValidatedProjection Validate(RenderedChartRecord record)`
- `src/Chart.Renderers/NativeRenderingCore.cs:91` — `public static string Encode(RenderedChartRecord record)`
- `src/Chart.Renderers/NativeRenderingCore.cs:109` — `public static RenderedChartRecord Decode(string text)`
- `src/Chart.Renderers/NativeRenderingCore.cs:156` — `internal NativeRenderException(string operation, bool rollbackComplete, Exception inner) : base("Native rendering failed at " + operation + "; rollback=" + rollbackComplete, inner)`
- `src/Chart.Renderers/NativeRenderingCore.cs:162` — `public NativeSceneRenderer(IEditableSceneSurface surface)`
- `src/Chart.Renderers/NativeRenderingCore.cs:163` — `public NativeRenderOutcome Insert(string canonicalJson, string sceneJson)`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:26` — `public PowerPointSceneSurface(PowerPoint.Presentation presentation, PowerPoint.Slide slide)`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:28` — `public void VerifyAccess()`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:32` — `public void Begin(string operation, string chart)`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:54` — `public int CreateAnchor(ScenePlot plot, bool trailing)`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:60` — `public int CreateNode(SceneNode node, string token)`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:79` — `public int Group(IReadOnlyList<int> children)`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:94` — `public void Persist(int rootId, string chart, string xml)`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:104` — `public void Complete()`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:122` — `public bool Rollback()`
- `src/Chart.Renderers/PowerPointSceneSurface.cs:147` — `public RenderedChartRecord ReadVerified(PowerPoint.Shape root)`

## 7. Inputs and outputs

Canonical v2 JSON+sealed scene v1 -> NativeRenderOutcome rootID/record. Surface owns COM and receives validated IDs/DTOs.

## 8. Dependencies

ChartCodec/SceneCodec, XmlReader/Writer, PowerPoint and Office COM; owner STA/thread.

## 9. Workflow

Before Begin validate ID/revision/layout, ordered numeric bars/paint and omission diagnostics; cap2000 nodes/minimum 2 pt plot. Create two invisible anchors; omit zero-area native shape but retain binding. Group only operation-owned children, persist verified CustomerData XML, complete. Any mutation exception rolls back owned objects and reports rollback status.

## 10. Integration points

Canonical v2 JSON+sealed scene v1 -> NativeRenderOutcome rootID/record. Surface owns COM and receives validated IDs/DTOs. Compiled bounded feasibility path; native runtime evidence remains separate from unit/source proof.

## 11. Exceptions and recovery

NativeRenderException has Operation and RollbackComplete. ReadVerified rejects missing/ambiguous parts, duplicate chart roots/tokens, unexpected nodes and geometry/style drift; never rewrites data to fit edits.

## 12. Configuration dependencies

2000-node feasibility cap; 0.05pt geometry and0.001 transparency read tolerance; rectangle-only surface; XML cap12,582,912 chars.

## 13. Security considerations

Operation/chart/node tags and ownership tree prevent deleting user shapes during rollback. XML DTD prohibited/resolver null; bindings require stable SHA token and positive unique shape ID for area nodes.

## 14. Unit and integration test definitions

- `tests/s3/NativeRendererTests.cs:1`
- `tests/s3/RendererRibbonTests.cs:1`

Inspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.

## 15. Known limitations

No axes/legend output, native edit-to-canonical reconciliation or general family adapter. Manual edits survive as native content but metadata verification detects drift. Duplicate IDs fail closed.

## 16. Technical debt

Surface owns operation COM references; full production command/revision/approval boundary remains required by source comment. General geometry and SLO acceptance absent.

## 17. Extension points

Expand only with capability/version tests and inverse mapping; retain preflight-before-mutation and ownership rollback.
