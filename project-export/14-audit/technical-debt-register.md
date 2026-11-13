# Technical debt register

These are source-review findings, not a remediation commit or a fresh Checkmarx result. Severity is qualitative maintenance/security impact; no exploit is asserted unless demonstrated.

| ID / priority | Finding and evidence | Consequence | Recommended engineering response |
|---|---|---|---|
| TD-01 High | V3 chart classes exist but both active projects omit them (`src/Chart.Domain/Chart.Domain.csproj:14`; `apps/powerpoint-vsto/PowerPointIntelligence.csproj:45`) | Engineers can mistake authored options for persisted production capability | Deliberate compile/migration decision, direct V3 contract/compatibility tests; do not silently add to build |
| TD-02 High | Active source/imported-original/procedural520 modes coexist (`apps/powerpoint-vsto/TemplateLibrary.cs:1`; `src/Library.Core/PresentationSource.cs:1`) | Library counts/design provenance/editing semantics can be conflated | Keep provenance/capability UI and tests explicit; preserve license boundaries |
| TD-03 High | TestS1 fallback NuGet download before independent identity verification (`scripts/Test-S1.ps1:18`) | Supply-chain coverage differs across test/build paths | Reuse a pinned approved digest/signature path, fail closed and tamper-test |
| TD-04 High | Scanner collector targets scene/template runners, not PIIDX (`scripts/Collect-SceneThreatEvidence.ps1:23`) | Historical absence of a PIIDX match was insufficient evidence | Scope/positive-control actual PIIDX file/hash without restoring detected binaries |
| TD-05 Medium | Locale-dependent fingerprint ordering (`scripts/Source-Fingerprint.ps1:1`; `easySetup/Common.ps1:14`) | NLS/ICU can mark unchanged package stale | Ordinal shared canonical recipe plus PS5/PS7 vectors and new signed release |
| TD-06 Medium | Broad UI catch and dense single-line implementations (`apps/powerpoint-vsto/ThisAddIn.cs:133`) | Root cause visibility and safe review become difficult | Keep user-safe status and improve fixed structured diagnostics; focused refactors with native regression |
| TD-07 Medium | Original native adapter is narrower than scene DTO/TS engine (`src/Chart.Renderers/NativeRenderingCore.cs:18`) | Type presence can imply unsupported native geometry | Explicit capability matrix and preflight rejection; add families one tested path at a time |
| TD-08 Medium | MJS utility calculations lack consistent result/accumulation finite guards (`packages/chart-core/src/calculations.mjs:1`) | Extreme finite inputs can overflow offline results | Negative overflow tests and explicit error/result policy; no production native impact assumed |
| TD-09 Medium | Source reader bounds entries/XML but no aggregate decompression budget (`src/Library.Core/PresentationSource.cs:1`) | Large packages can still consume work across entries | Aggregate byte/count/time/cancellation limits and stress fixtures |
| TD-10 Medium | Related installer/package readers retain implicit XML casts (`scripts/Install-Local.ps1:21`; `scripts/Test-DemoPackage.ps1:27`) | Explicit safe XML policy is inconsistent | Consolidate bounded DTD-prohibited readers; prove no app/config regression |
| TD-11 Medium | Native tests are environment/foreground/token dependent (`tests/s0/ContextPropertiesOfficeTests.cs:1`) | Infrastructure prerequisites can be mislabeled feature failure/pass | Preserve BLOCKED distinction, owned cleanup and per-feature evidence |
| TD-12 Medium | Enterprise requirements and current starter README differ from implemented desktop breadth (`README.md:1`; `services/ai-gateway/README.md:1`) | Onboarding/acceptance confusion | Maintain source-backed traceability and separate roadmap from shipping subsets |

Duplicated root chart-core.yml is not an active workflow; prefer one authoritative CI definition. Symbolic declaration extraction has known limitations with compressed statements; use source bodies for consequential edits. No claim of exhaustive dead-code elimination or automated call-graph reachability is made.

## Test-oracle and cleanup follow-up

- TD-13 Medium: `tests/s0/NativeFormattingTests.cs:1` cancel assertion uses an unchanged `touched` variable; add an oracle that would fail against a known mutating implementation rather than a constant. No production mutation defect is inferred from that vacuous assertion.
- TD-14 Medium: `tests/s0/MappedExcelOfficeTests.cs:1` closes its workbook/releases RCW without explicit Quit for its separately created Excel application; `tests/s0/DataLinkManagerTests.cs:1` leaves temp fixtures. Track owned processes/directories and verify cleanup without touching user Office sessions.
- TD-15 Medium: `tests/s0/DemoContractTests.cs:1` literal scenario-presence checks are structural traceability, not per-feature runtime tests. Preserve that distinction in release acceptance and strengthen focused actual behavior coverage.
