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
- S0-AT-001: Add-in loads in target PowerPoint and ribbon command executes.
- S0-AT-002: Inserted object is editable through intended host controls.
- S0-AT-003: Save/reopen restores chart ID, data, and formatting with no corruption.
- S0-AT-004: Duplicated slide has no chart identity collision.
- S0-AT-005: Missing/corrupt metadata yields safe diagnostics, not data loss.

**Current status:** IN PROGRESS. Implementation and test harness added; no S0 acceptance criterion is marked passed. See [execution evidence](../evidence/S0-2026-10-21.md) and [Windows validation runbook](S0-windows-validation.md). VSTO build/load and actual PowerPoint execution remain unverified. **S1 is gated.**
