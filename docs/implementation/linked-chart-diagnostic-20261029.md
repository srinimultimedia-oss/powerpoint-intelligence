# Focused linked-chart diagnostic — 29 October 2026

Owner run: installed 1.0.0.9 / piidx7. All element/agenda suites pass. Workbook write read-back did not report a failure, but native chart series still reports zero points during mapped refresh; PowerPoint later becomes RPC-unavailable. This does not establish whether the root cause is series binding/cache, workbook close handoff, Office build behavior or an external component. No new speculative binding fix was applied in this increment.

A focused mode is now available:

```powershell
.\scripts\Test-DemoFeatures.ps1 -InteractiveOffice -LinkedChartDiagnostic
```

No reinstall is required for this diagnostic. It builds the existing test runner at its existing path, creates one owned CSV fixture and one chart in a temporary presentation, and calls the current write path. It records metadata before binding, after worksheet verification, after binding, before close and after close. Recorded fields include Office file/product version, OS/runtime/bitness, series counts, raw array type/rank/lower bound, normalized lengths, reference-error/external-workbook flags and whether ChartData is linked. No source values, series captions or formula text are logged.

When possible it preserves the generated test presentation as linked-chart-success.pptx or linked-chart-failure.pptx so chart caches and embedded workbook relationships can be inspected. These artifacts contain only the known generated fixture. It closes only its owned presentation and never quits the user's PowerPoint process.

The script also reads Application Error events from the preceding 30 minutes and records only timestamp, POWERPNT/EXCEL name, faulting module and exception code. Missing event permissions are recorded as unavailable; they do not become crash-free evidence.

The optional thread-local write observer defaults to null in normal operations and is reset by the probe's finally block. The probe is diagnostic coverage of one initial chart write, not proof of refresh, every chart family, native formatting or save/reopen acceptance. Existing business behavior, chart timeout and failure conditions are unchanged.

Verified locally: focused runner compilation, script syntax, existing ribbon/element contract suite and git diff whitespace check. Native execution remains unavailable under this agent's AppContainer token; desktop evidence is required. Read linked-chart-stages.tsv, the generated fixture if present and office-crashes.json together before selecting another product fix.