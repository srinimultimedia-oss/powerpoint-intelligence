# Technical handover

## First orientation

Read the [master design](MASTER-TECHNICAL-DESIGN.md), [module inventory](../03-modules/module-inventory.md), [actual feature status](../05-features/feature-status-matrix.md), [current test evidence](../12-testing/test-results.md) and [security gaps](../09-security/security-gap-register.md). Match commit and per-file hashes before using line references. Visible brand is NexWave; existing PowerPointIntelligence COM/metadata/storage identities remain compatibility contracts.

## Runtime maintenance map

ThisAddIn owns startup, callbacks, UI gating and cleanup. PowerPointHost owns S0 chart/metadata association and UI STA checks. DataSources and mapped-link helpers own bounded input/provenance/commit observation. Element modules own text/process/connector/indicator/agenda state. RendererRibbonCommand/native surface owns the narrower grouped-shape projection. Gallery/workspace/Library.Core own local registration/extraction/search/preview/import. Module maps give all source assignments and method locators.

## Build and change discipline

Use VisualStudio Office tools/Framework4.8 and the approved signer in normal Windows PowerShell. Test deterministic contracts first, then explicit native editing/save-reopen, and finally installed package/ribbon acceptance. Changing DTO definitions outside the compile graph is not a feature release; changing persisted namespace/schema requires explicit migration. Avoid broad clean/delete operations that remove installed extraction paths, stable receipts or security evidence.

## Immediate backlog

Address source-backed supply-chain/scanner-scope gaps and inconsistentXML policy; strengthen vacuous/structural test oracles and owned Office cleanup; define distribution publisher/Office matrix/data retention; make fingerprint ordering host-independent in a separately tested signed release. EnterpriseAI/auth/services need decisions and implementation before UI labels can promise live behavior. Do not import undocumented native/proprietary behavior or distribute licensed source assets beyond approved scope.

## Validation boundary

This export changes no application code and performs no install or dependency upgrade. Safe current checks cover229assertions, TS typecheck and50script parses. Office, live services, freshscanner clearance and production compatibility remain unverified here. Full-body/targeted review levels, exclusion reasons and diagram parser limitations are explicit. Use the issue registers as an implementation backlog, not completed remediation evidence.
