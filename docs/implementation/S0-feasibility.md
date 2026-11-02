# S0: PowerPoint host and persistence feasibility

## Goal
Demonstrate a Windows PowerPoint add-in can create an editable chart, persist canonical metadata, save, reopen, and recover the chart.

## Prerequisites
Windows + desktop PowerPoint, Visual Studio with Office development tools, .NET Framework 4.8 developer pack, a local test certificate/signing approach, test PPTX corpus.

## Tasks
1. Record Office product/build/bitness and machine prerequisites.
2. Create PowerPoint VSTO Add-in project and a Ribbon with `Insert Sample Chart`.
3. Add task pane displaying selected chart ID and status.
4. Implement `IPowerPointHost` from `apps/powerpoint-vsto/Contracts.cs`.
5. Insert an editable chart with three categories and two series, or native shapes with explicit limitation.
6. Serialize `ChartDocument` metadata using a supported persistence mechanism and associate with native objects.
7. Save, close PowerPoint, reopen, load metadata, edit one data value, and save again.
8. Test duplicate slide, chart copy to a new presentation, deleted objects, and malformed metadata.
9. Capture objective evidence and record gaps. Avoid reliance on unstable shape IDs without recovery strategy.

## Exit criteria

Cross-presentation copy limitation (22 October 2026): owner runtime evidence
shows native Shape.Copy/Shapes.Paste can retain the chart tag without its
CustomerData XML. The explicit adapter operation
`CopyEditableChartToActiveSlide` transfers and verifies metadata alongside the
native chart; its interactive test remains pending. Ordinary cross-document
clipboard metadata preservation is not supported or marked passed. No ribbon
command is wired to this adapter method yet.
- S0-AT-001: Add-in loads in target PowerPoint and ribbon command executes.
- S0-AT-002: Inserted object is editable through intended host controls.
- S0-AT-003: Save/reopen restores chart ID, data, and formatting with no corruption.
- S0-AT-004: Duplicated slide has no chart identity collision.
- S0-AT-005: Missing/corrupt metadata yields safe diagnostics, not data loss.

**Current status:** ACCEPTED BY OWNER in chat on 21 October 2026, followed by explicit authorization to start S1. No S0 acceptance test is marked passed without execution evidence. See [execution evidence](../evidence/S0-2026-10-21.md) and [Windows validation runbook](S0-windows-validation.md). VSTO build/load and actual PowerPoint execution remain unverified; owner acceptance authorizes progression with those limitations recorded.

### Later direct desktop UI evidence — 22 October 2026

The earlier statement that all VSTO loading/PowerPoint execution was unverified is superseded narrowly for the previously installed package (receipt 1.0.0.3). A supported native UI test opened PowerPoint 16.0.20527.20088 x64, invoked its Intelligence > Elements > Managed sample column command, edited the embedded native workbook, saved, closed and reopened an owned deck. The original chart GUID and revision 1 were shown after reopening; closed-package checks matched the edit across XML metadata, chart cache and XLSX. See artifacts/activation-diagnostics/legacy-ui-persistence.json and artifacts/ui-evidence/legacy-native-chart-reopened.png. This supports S0-AT-001/002/003 for that one installed legacy fixture only. It does not accept the current source build, the new grouped renderer, S0-AT-004/005 or untested Office editions. Scripted COM activation from this restricted agent context still returns 0x80070005.

## Desktop adapter validation — 26 October 2026

The owner's `artifacts/validation-20261026-155936-44d22bed` run passes office-1, office-2 and office-3 adapter integration, installed ribbon identity/assembly hash/OnLoad, and native-shapes-office. This verifies the executed adapter insertion/edit/persistence/recovery assertions under desktop x64 Office. It does not by itself verify full process restart or task-pane visual behavior. The separate original-template comparator and demo report remain under investigation as documented in thinkcell-observable-study.md; they are not three new S0 adapter failures.
