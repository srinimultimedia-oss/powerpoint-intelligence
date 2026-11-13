# Testing architecture

## Layers and test seams

The repository uses custom C# Console assertions, Node's built-in `node:test` and Python `unittest`; no centralized instrumented coverage tool is configured. Host-free tests exercise codecs, validation, calculations, scene hashing, rollback using fake surfaces, library extraction/schema/cache policies and formatting planning. UI/COM tests are opt-in and require an interactive desktop with valid Office/VSTO activation.

`tests/s3/NativeRendererTests.cs:1` uses a fake `IEditableSceneSurface` with failure injection before/after mutation and checks owned rollback, reentrancy and owner thread. Its interactive mode separately checks real grouping, save/reopen, manual divergence and owned failure cleanup. A fake-surface pass does not prove Office native acceptance.

`tests/s0/HostIntegrationTests.cs:1` exercises embedded workbook edits, bounded native observation, metadata identity, source formatting, duplicate identity and source/destination copy. `ContextPropertiesOfficeTests` observes the installed COM API and distinguishes unavailable foreground focus from a panel mutation failure. No Windows focus policy is changed by those prerequisites.

Library tests have synthetic ZIP/PNG and schema corruption fixtures plus optional Office fidelity runs. Synthetic preview-cache acceptance does not establish native rendering; source-original fidelity combines bounded raster comparison with exact geometry/formatting and selected edit/save/reopen evidence. External licensed corpus fixtures are environmental inputs, not distributed test assets.

## Orchestration

`scripts/Test-All.ps1:1` creates a source snapshot and output under artifacts, preserves stable runner caches via RunnerRoot, runs host-independent checks, architecture builds and optional installed/interactive tests. It validates the snapshot against source fingerprints and uses owned-process timeouts. This export did **not** invoke it, because its filesystem writes violate the export-only assignment. Existing test scripts often compile in artifacts or use normal temp folders; they are documented, not silently redirected with changed application behavior.

CI: four active workflows run portable calculations, Office compatibility configuration, S0 metadata and S1 contracts/schema checks. CI does not load PowerPoint/VSTO or prove enterprise deployment. Root chart-core.yml is a duplicate design file outside .github/workflows, not an additional GitHub Actions workflow.

## Current constrained validation

The export invokes unchanged read-only Node test modules, TypeScript --noEmit, script AST parsing and a direct C# metadata compilation with output and temp under project-export. No existing binaries or receipts are replaced. Logs and exact Node argument arrays/times are stored beside this documentation. Source-preservation verification provides a separate gate from test success.
