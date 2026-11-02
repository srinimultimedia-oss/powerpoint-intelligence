# Embedded-workbook handoff observation — 29 October 2026

Evidence: artifacts/demo-features-5543965d3aa3463b9ed1f850e9748070/linked-chart-stages.tsv. Actual Office file/product build is 16.0.20527.20088; process x64; CLR 4.0.30319.42000. The reported OS 6.2 in an unmanifested .NET process is not treated as authoritative Windows-version evidence.

The generated chart starts with three sample series/four points. After worksheet verification and SetSourceData, it correctly has two series/three points each, no reference-error flag and no external-workbook formula flag. Immediately after workbook close it still has two series and the first series returns three points. The following diagnostic read of Series.Formula fails with RPC_E_CALL_FAILED; subsequent COM reads report RPC server unavailable. The second post-close series was not read. This narrows the transition but does not prove whether the trigger is workbook close, the diagnostic read, installed callbacks or Office internals.

No product write change is made from this observation. The focused probe now records only the close-return timestamp at the immediate handoff, without any post-close COM snapshot; formula classification is limited to the source-bound stage while the workbook is open.

Baseline:

```powershell
.\scripts\Test-DemoFeatures.ps1 -InteractiveOffice -LinkedChartDiagnostic
```

If it still fails, restart PowerPoint with presentations closed and compare:

```powershell
.\scripts\Test-DemoFeatures.ps1 -InteractiveOffice -LinkedChartDiagnostic -PauseOwnAddIn
```

The optional comparison identifies exactly one PowerPoint Intelligence COM add-in, disconnects only it, runs the same compiled write path and restores its original connected state in finally. No third-party/security add-in, antivirus setting, document policy or COM security boundary is modified. If PowerPoint becomes unavailable, restoration may fail; the diagnostic reports this explicitly and the owner may need to re-enable PowerPoint Intelligence through COM Add-ins after restart. Other add-ins remain untouched, so this does not isolate them.

Crash collection now distinguishes an empty event query from an actual read failure rather than reporting every no-match as unavailable. Only application/module/exception metadata is saved. Neither baseline nor comparison is a full native acceptance run or a false-positive verdict.

Checks: focused runner compilation passed; script syntax and diagnostic-only flag guard passed without starting Office. No reinstall is required for this diagnostic-only refinement. Native execution of these comparisons requires the normal desktop token; this agent's existing AppContainer activation restriction remains unchanged.
## Desktop comparison follow-up

Both desktop runs `demo-features-1c6d43387b624d329843b3298fc40bfd` and `demo-features-46dd7791405e4c6c8a6be446c93d8225` reached write-returned. Independent ZIP inspection of both saved generated fixtures confirms two native chart series, three cached points per series and one embedded XLSX. The paused run records own-addin-restored=True. These establish initial-write success with and without this add-in; they do not establish a root cause or full linked-chart acceptance. Instrumentation timing remains a possible confounder.

The focused diagnostic now removes its observer before refreshing a changed synthetic CSV, edits the native chart title, saves/closes/reopens its owned presentation and compares all six expected numeric values plus the title. It never changes user presentations or security policy. `Test-DemoFeatures.ps1` compiled this revised runner successfully; Office execution was skipped in the agent session. Desktop execution remains required using `Test-DemoFeatures.ps1 -InteractiveOffice -LinkedChartDiagnostic`. No add-in reinstall is required for this diagnostic-only change. Full mapped Excel and feature-suite acceptance remains open.

## Refresh failure and readiness gate

Desktop run `artifacts/demo-features-071decbd34284f19ae2c3492553e2e08` passed its instrumented initial write and failed the uninstrumented refresh: expected three points, actual zero. ZIP inspection of the saved failure fixture confirms twelve embedded worksheet cells, two local series range formulas and zero cached value points in both series. This establishes a chart/workbook disagreement; it does not establish an Office crash or a complete root cause.

`LinkedChartData.Write` now checks the actual expected native chart values, with the existing eight-second bound, before closing the embedded workbook. It retains the post-close verification and has no sample fallback. The new stage is chart-values-acknowledged. This is a candidate handoff fix, pending desktop evidence; it does not change security controls. The standalone diagnostic compiles and exercises the current source without reinstalling the add-in. Run `scripts/Test-DemoFeatures.ps1 -InteractiveOffice -LinkedChartDiagnostic` to test initial binding, uninstrumented refresh, native title editing and save/reopen. Installation and broader native acceptance remain pending.
