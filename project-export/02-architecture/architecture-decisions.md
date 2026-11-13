# Implemented decisions and unresolved architecture

The controlling decision register is `docs/DECISIONS.md:1`. This export does not update or silently resolve it.

| Decision family | Observed implementation | Boundary |
|---|---|---|
| AD-001/006 Windows VSTO | Framework 4.8 C# add-in and Office COM adapters | Actual estate/version certification is not established by a profile string |
| AD-004/009 native editability/metadata | Native Office charts/shapes and CustomXMLPart plus tags | Metadata validity is checked; native edits can create detectable divergence |
| AD-008 deterministic TypeScript | Six basic engine families, strict scene contracts, offline CLI | Not a complete native renderer for all required chart families |
| AD-013 grouped shapes | Native scene surface and original template compositions | Grouped shapes are editable drawing objects, not native Excel-backed chart types |
| AD-014/017 explicit linked data | Local source selection and confirmed worksheet/range mapping | Changed headers/missing files stop refresh; no enterprise scheduling |
| AD-015/016/019 libraries | Procedural 520-layout catalog and local original-file source library | Different content modes; licensed source reuse is not independently authored parity |
| AD-018 raster tolerance | Bounded edge-local antialias comparison plus native geometry/format checks | Sixteen original slides are a sampled historical fidelity test, not all 497 |
| AD-020/021 setup and brand | Dual-bitness self-contained VSTO setup and NexWave display name | Compatibility IDs/storage names remain PowerPointIntelligence |
| AD-002/003/007 on-prem and AI | Service schema/profile scaffolds | Production service, Azure invocation, RBAC/SSO remain unimplemented |

Open Q-001..Q-012 include exact tested Office builds, identity, Azure resource/networking/data policies, enterprise infrastructure, signing distribution trust, licensing and performance/locale policy. Local SQLite implementation does not resolve enterprise Q-006. Configurable URLs do not resolve auth/data-governance Q-003..Q-005. Historical owner S0 acceptance is a project decision, not fresh runtime evidence in this export.

Design tension: active procedural/native/original-library paths and legacy S0/v2/V3 contracts coexist. Extension work must choose a production representation and codec intentionally, preserve source compatibility identities and add explicit migration/round-trip tests before claiming a new persisted version. `src/Chart.Domain/Chart.Domain.csproj:12`; `src/Chart.Domain/S0Migration.cs:1`.
