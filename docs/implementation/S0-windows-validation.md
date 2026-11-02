# S0 Windows validation runbook

Use an interactive Windows development desktop with an approved PowerPoint build. Do not run Office automation on a server or hosted CI agent.

## Build and checks

1. Record machine identifier (locally), Windows, PowerPoint File → Account → About version/build/bitness, .NET Framework targeting pack and Visual Studio version. Confirm Q-001/Q-002/Q-007 with owner. Install missing Office development workload/4.8 developer pack only through approved machine administration.
2. Open `apps/powerpoint-vsto/PowerPointIntelligence.csproj` in Visual Studio with Office development tools. Build for .NET Framework 4.8. Validate the hand-authored host bootstrap against the VSTO tooling; do not call the project build verified until it succeeds. Select an approved local test certificate through project signing settings or pass `/p:ManifestKeyFile=<local-test.pfx>`. Never commit certificates or passwords. Use the Visual Studio debugger's VSTO deployment workflow; production distribution remains undecided.
3. Run `./scripts/Test-S0.ps1` from PowerShell on that machine (or supply `-Compiler <Roslyn-csc.exe>`). This compiles/runs the 17 serializer tests and compiles the COM integration executable. If application policy blocks the executable, obtain an approved execution environment from the administrator; do not bypass policy.
4. Run `./scripts/Test-S0.ps1 -InteractiveOffice` explicitly in the desktop session. Preserve stdout/stderr and generated artifacts/office-*/roundtrip.pptx and copied.pptx. This opens visible PowerPoint and checks actual native charts, data, formatting, duplicate/copy, missing/corrupt XML, deleted objects and a second edit/save/reopen. It does not prove VSTO loading or process restart. Keep failed fixtures for investigation; the test never quits the user's PowerPoint process.
5. Run `npm test` in packages/chart-core and `git diff --check` at root. Metadata CI is a host-independent check only.

## Manual VSTO acceptance

1. With the add-in loaded in PowerPoint, create a blank slide. Capture the Intelligence ribbon and task pane in a screenshot; click **Insert Sample Chart**. Record chart ID from pane (AT-001).
2. Verify **Chart Design → Edit Data** exposes Q1/Q2/Q3, Actual 10/20/30 and Plan 15/18/35. Change Actual/Q1 to 42; set chart style/color and resize using native controls. Click Inspect / Recover; confirm selection ID and revision/status. Record numeric values and screenshots (AT-002).
3. Save as PPTX, close all test decks, exit PowerPoint normally, confirm POWERPNT.EXE has exited, then restart and reopen the file. Verify the same chart ID, 42 and formatting. Edit another data value, save, fully restart again and recheck (AT-003). Do not kill processes or close unrelated user work.
4. Duplicate the slide, click Inspect / Recover, compare each selected chart's ID and data. IDs must be distinct before save and after full restart. Repeat with copying a chart to a new presentation; XML and data must travel, with no collision after additional copies in that deck (AT-004).
5. On disposable copies only, use the integration harness corruption/missing-part cases. Inspect each damaged chart in the loaded add-in: status must identify diagnostics and the native chart must remain editable. Do not overwrite the only good deck. Delete a chart, save/reopen and confirm other charts still resolve correctly (AT-005).
6. Also exercise two open decks, selection switching, no document, slide sorter/no slide selection, read-only deck, cancel save, and shutdown. Check task pane is associated with the correct window; document any gaps.

Attach logs/screenshots with machine and Office build to the evidence report. Mark each criterion passed only for an actual successful run. If any criterion fails or cannot run, S0 stays open and S1 stays gated.
