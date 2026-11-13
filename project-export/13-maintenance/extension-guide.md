# Extension guide

## Add a chart or visual family

Choose among Excel-backed native chart creation, grouped editable shapes, procedural library layouts and faithful original slide insertion. These routes have different persistence/editability semantics. Define model constraints and capability rejection first. Add pure calculations/scene tests before Office mutations, then map native semantic bindings and preserve omission diagnostics. `src/Chart.Renderers/NativeRenderingCore.cs:158`; `src/Chart.Renderers/PowerPointSceneSurface.cs:14`.

The native feasibility adapter verifies a canonical v2 clustered-column model against a rectangle scene. Extending TypeScript `buildScene` to another family alone does not extend that C# adapter. Do not silently reinterpret V3 files that are not compiled into current targets. `src/Chart.Domain/Chart.Domain.csproj:12`; `packages/chart-engine/src/scene-builder.ts:13`.

## Add UI commands

Add a catalog/definition entry, callback route, enablement policy and actual implementation. Use the existing UI STA, guarded `Run`, selection identity and undo boundary. Include tests asserting the callback exists, the target is appropriate and unsupported selections fail before mutation. Keep manually chosen native formatting when canonical refresh policy requires it. `apps/powerpoint-vsto/ThisAddIn.cs:133`; `apps/powerpoint-vsto/DemoRibbon.cs:1`.

## Extend local library/data integration

PIIDX changes require explicit schema version and migration compatibility, parameterized statements, source fingerprint refresh, tested duplicate/search behavior and preview integrity. A new input source needs bounded parsing, explicit mapping and a source-change failure policy. Never treat a local profile as an authenticated connector. `src/Library.Core/PiidxSchema.cs:1`; `src/Library.Core/PiidxCatalog.cs:1`; `apps/powerpoint-vsto/DataSources.cs:1`.

## Add governed AI/services

Implement approved identity/network/data-policy decisions, transport client, request/response schemas, timeout/retry policy and action approval before wiring a ribbon command. Existing JSON schema is a proposed contract, not a running orchestration engine. Do not permit model text to invoke COM/code/SQL. `services/ai-gateway/README.md:1`.
