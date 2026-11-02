# Native mutation coordination — 29 October 2026

Desktop evidence: installed 1.0.0.5 / piidx3; `artifacts/demo-features-b23a46e70dc44ae9899a5e5b25657277`.

Verified progress in that run: contextual properties 23, font 8, productivity 13 and indicator 16 native assertions pass. Styled grid insertion and source metadata independence pass. Text fitting/conversion passes until outline color. These are owner desktop results for that payload, not new native passes from this session.

## Diagnosis

`connector-saved-state.txt` shows Saved=false at before-enumeration and planned-changes-0. Therefore that run does not show reflow itself dirtying a previously clean presentation. The agenda reports one registry whose identity and payload both differ from the captured part. The installed add-in's process-local busy flag does not coordinate its automatic writes with the separate compiled Office test process. Competing writes are a supported explanation; the new coordination must still be proven in Office.

Grid reorder also had an independent defect: geometry preflight called managed ConvertSelected/Configure on a temporary duplicate. Those methods persist XML, while the duplicate can share the original attachment. Preflight now changes only temporary native geometry and measures it; it never persists metadata on the probe.

## Changes

- A per-PowerPoint-process named mutation mutex coordinates this add-in's commands/automatic callbacks with explicitly cooperating desktop automation. Native tests acquire the lease before creating their fixtures and dispose it across every return/exception. Installed automatic writes defer while the external test owns it; context-property read observation remains enabled. Normal add-in commands retain their own busy guard, and kernel ownership is released if the test process exits. No security permissions or Windows policies are changed.
- Five pure contention/release assertions cover acquire, bounded competing access, exclusion, same-thread nesting and release. These test mutex semantics, not actual cross-process Office acceptance.
- Outline color sets visibility before RGB, avoiding hidden-line initialization overwriting the requested color. Native outcome remains a gate.
- Read-only process-grid geometry preflight replaces mutation/persistence on temporary probes.
- Workflow cleanup retains and reports its original exception rather than replacing it with a later RPC error from cleanup.
- Command summaries count commands actually executed and blocked. After an Office disconnection, unexecuted commands remain BLOCKED and cannot be described as tested.

New payload marker: `PPI-DEMO-RIBBON-20261029-v17-piidx4-mutation-coordination`.

## Checks and limits

Local full suite: `artifacts/validation-20261029-163746-8d88b29f/results.json`, 23/23 PASS. That snapshot precedes the final read-only grid preflight patch. After that patch, the ribbon/element contracts and native-test compilation were rerun, plus all production architecture builds; logs are in `artifacts/native-concurrency-20261029`. XML/locale/coordination contracts total 38 assertions.

Native PowerPoint activation remains unavailable to this agent's AppContainer token. The new payload needs desktop installation and `Test-DemoFeatures.ps1 -InteractiveOffice`. The coordinated test run validates explicit native operations, not automatic timer delivery; automatic update behavior without the test lease requires separate acceptance. Third-party add-ins or uncoordinated external automation are not serialized by this protocol. No native failure is marked fixed solely because local checks pass.