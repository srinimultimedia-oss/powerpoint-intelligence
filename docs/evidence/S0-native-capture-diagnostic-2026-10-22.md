# S0 native capture investigation — 22 October 2026

Owner-provided desktop execution: Office version 16.0, x64 test process. All 17
metadata tests passed. Native clustered-column insertion, two series, three
categories and attached metadata passed. Native workbook edit reconciliation
reported zero healthy charts and one diagnostic; the edit assertion failed.
Save/reopen, duplication and later recovery assertions were not reached.

The existing catch discarded the reason. Added an opt-in diagnostic callback
with fixed operation names, exception type and HRESULT, excluding exception
messages, IDs, cell values and metadata contents. The interactive runner prints
these diagnostics and checks their presence for its corrupted-metadata fixture.
No speculative capture behavior change has been made.

Local verification: `scripts/Test-S0.ps1` passed all 17 metadata tests and compiled
the updated Office integration runner. Interactive integration was skipped in
this execution context. The new corrupted-metadata diagnostic assertion is
compiled but not yet executed. Reproduce from the owner's desktop with
`scripts/Test-S0.ps1 -InteractiveOffice` and retain the operation/error line.

S0 native edit capture and persistence acceptance remain open. No S1 expansion
is included in this change.

## Follow-up: series access failure

Owner rerun isolated `capture.series-values`, `InvalidCastException`, HRESULT
`0x80004002`. This operation previously combined dynamic series lookup, property
access and an Array cast, so the evidence does not identify which cast failed.
Replaced dynamic chart/series access in Capture with typed PowerPoint interfaces.
Explicitly materialize Values/XValues as object before normalizing their VARIANT
payload, avoiding dynamic dispatch of the normalization call. Support rank-one
typed/nonzero-bound SAFEARRAYs and scalar single-point values; reject higher ranks.
Separate series-item and series-values diagnostic stages remain available.

Verification: `scripts/Test-S0.ps1` passed 17 metadata tests and six new series
normalization tests; updated integration runner compiled. Office integration was
skipped in this restricted session. The owner's interactive rerun is required to
verify the fix and continue through persistence/duplication acceptance.

## Follow-up: native clipboard transfer loses attached XML

Full owner-provided output confirms PASS for native workbook edit capture,
save/close/reopen identity and data, formatting persistence, duplicate-slide ID
repair, and duplicate metadata retention. The failure occurs in the destination
presentation after raw Shape.Copy/Shapes.Paste: the ID tag survives but Read finds
no matching CustomerData metadata. This is runtime evidence on reported Office
16.0 with an x64 test process; no other Office version is certified by it.

Added explicit `CopyEditableChartToActiveSlide` adapter operation. It validates
and captures the source, copies the native chart, writes and verifies the XML on
the pasted shape, and assigns a new chart identity if the destination already
contains the ID. Failed metadata transfer removes only the newly pasted shapes.
No automatic lookup or resurrection of missing metadata from other open files.
The integration test now explicitly tests adapter transfer, native formatting,
and destination save/reopen; it no longer claims raw clipboard metadata support.

Local verification: 17 metadata plus six conversion tests pass; integration
runner compiles. Interactive adapter transfer and subsequent corruption recovery
checks require the owner's desktop rerun. Native Ctrl+C/Ctrl+V across documents
remains unsupported for canonical metadata and is not an acceptance pass.

## Follow-up: intermittent stale chart values

The next owner run reported one healthy chart and zero diagnostics but failed
native-edit capture before reaching copy. Refresh timing is a hypothesis, not
an established cause. The runner now checks the fixture cell write, calculates
only that worksheet, closes the workbook in finally, then pumps its STA message
queue and polls native series values with Refresh for at most ten seconds.
A separate precondition failure identifies workbook/chart synchronization before
capture. Metadata capture is not retried and expected values are unchanged.
This synchronization change is test-only; production capture is unchanged.
Local checks: 23 non-Office tests pass; integration runner compiles. Desktop
execution is required to establish the cause and validate the remaining checks.

## Follow-up: embedded Excel RPC disconnection

Owner rerun passed the embedded cell write assertion, then failed with
0x800706BA at a dynamic COM call taking one Boolean argument, consistent with
workbook.Close(true). The preceding chart Refresh ran before that Close.
Removed the test's pre-close host refresh, and moved the adapter SetData refresh
after workbook.Close(true). This is a targeted sequencing correction; refresh
invalidating the proxy remains a hypothesis, not a proven root cause.
The runner now emits fixed calculate/close/wait operation labels and compiles
with PDB symbols for source-line stacks. No RPC error is suppressed or retried.
Local check: all 23 non-Office tests pass; updated adapter/test runner compiles.
Interactive RPC stability and full S0 acceptance remain unverified.

## Follow-up: source presentation context

Owner run passed workbook synchronization, capture, save/reopen, formatting and
slide duplication. The line-numbered failure occurs at Read(source) inside the
explicit copy operation, BEFORE clipboard copy or paste. The destination is
active then. This weakens the earlier inference that native clipboard transfer
itself dropped XML: inactive-document CustomerData access is now implicated;
raw clipboard metadata behavior remains unproven.

The adapter now retains the destination window/slide, activates the source
window/slide for Read, Capture and Copy, and restores the destination in finally
before checking identity collisions and pasting/writing. The integration runner
also verifies source metadata after returning from transfer. Local 23 tests
pass and adapter compilation succeeds; interactive context handling is pending.
No full S0 acceptance or cross-version certification is claimed.

## Follow-up: duplicate replacement ownership

Owner rerun still fails Read(source) after source activation. Active-document
context alone is therefore not a sufficient explanation. The previous test
checked only the duplicate after recovery, not the original. Added an assertion
for original metadata immediately after duplicate recovery and before creating
the destination presentation, to distinguish replacement loss from document
switching. No root cause is marked established yet.

Write now caches the prior part ID before Add and uses shape.CustomerData.Delete
for replacement/rollback instead of document-wide CustomXMLPart.Delete. The
corruption fixture likewise removes parts through the shape collection using
IDs collected before deletion. Microsoft documents CustomerData.Delete as
removing the specified part from the object's collection, and CustomXMLPart.Delete
as deleting it from the XML data store:
https://learn.microsoft.com/en-us/office/vba/api/powerpoint.customerdata.delete
https://learn.microsoft.com/en-us/office/vba/api/office.customxmlpart.delete

Local verification: all 23 non-Office tests pass and integration runner compiles.
Office regression checks require desktop execution; this change is not yet a
verified runtime fix. The prior clipboard-loss inference remains unconfirmed.

## Follow-up: explicit RPC call rejection at workbook close

Owner stack now identifies workbook.Close(true), HRESULT 0x80010001
RPC_E_CALL_REJECTED. This run did not reach duplicate metadata regression checks.
Added a shared bounded retry helper used only for workbook Close in the adapter
and fixture: at most 51 attempts, 100 ms between explicit call rejections (five
seconds of retry delay, excluding COM call duration). Disconnection 0x800706BA
and all other failures propagate immediately; no mutation is broadly replayed.
The standalone STA runner pumps messages between close retries. Production code
waits on its STA without introducing a DoEvents reentrancy loop.
Microsoft documents time-bounded handling of rejected Office COM calls:
https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office

Local verification: 17 metadata plus ten series/retry assertions pass; adapter
and runner compile. Retry success, bounded failure and nonretryable disconnection
are tested with injected delegates without Office. Interactive stabilization
and full S0 acceptance remain pending; no new S1 work was undertaken.

## Comprehensive desktop run: validation-20261022-115051-3babf939

Read all saved results directly. All unit suites, cross-language validation and
three Compile targets passed. All three Office runs failed: runs one and two
lost the workbook connection on Close (0x80010108/0x800706BA); run three passed
capture/reopen/duplicate data but failed the NEW original-metadata assertion
immediately after duplicate recovery. This establishes loss during duplication
reconciliation before document switching, not during clipboard transfer.

Write now snapshots validated XML for other shapes referencing the same old
part ID before replacement. After removing that part, it restores a missing
other owner's XML as an independent attached part, or fails if existing XML has
changed. This preserves known valid data in the current presentation; it does
not invent missing data from tags or search unrelated documents. Interactive
shared-part regression remains pending.

Fixture edits now Save constant data before Close. A disconnected Close is not
replayed: the fixture reacquires a fresh workbook and asserts the saved value;
subsequent presentation reopen also checks the embedded workbook value itself.
This conditional verification is test-only; errors in reacquisition still fail.
No RPC disconnection is silently accepted as persistence success.

Corrected Test-All's Windows PowerShell stderr handling so exceptions do not
truncate child logs. The prior diff-check failure was a missing git.exe on the
desktop PATH (not whitespace warnings); use the standard Program Files Git
fallback and fail explicitly if unavailable. The runner prints its result table.

Updated local comprehensive run validation-20261022-115547-b42d643c passes all
12 noninteractive checks (201 unit executions, cross-language validation,
three architectures compiled, diff check). An intermediate wrapper-only run
115424 incorrectly reported failures due to locally shadowing LASTEXITCODE;
removed that assignment and verified the corrected runner exit status is zero.
Git fallback also supports the observed bundled Codex Git executable. Full
Office logs will be retained on desktop PowerShell instead of stopping on the
first stderr record. Shared-metadata preservation and fresh-workbook checks
are compiled but not yet validated interactively.

## Desktop run: validation-20261022-115713-780e5808

All 12 noninteractive checks passed. All three Office runs deterministically
failed at the newly added fixture Workbook.Save call with 0x800A03EC, "Unable
to get the Save property of the Workbook class". This was a regression introduced
by the test change, not evidence that the subsequent metadata repair failed.
Removed the standalone Save call. Persistence remains asserted independently
by presentation save/reopen followed by embedded workbook cell read, as well
as canonical metadata checks. No exception was broadly suppressed to pass it.

After removal, Test-S0.ps1 passes all 27 local assertions and compiles the runner.
Office execution is not available inside this session. The three-repeat desktop
run is still required to validate disconnection handling, original metadata
preservation, transfer and subsequent corruption recovery.

## Desktop run: validation-20261022-115956-970ff837

All three runs passed native edit capture; native embedded data and canonical
metadata across presentation save/reopen; formatting; duplicate identity repair
and BOTH original/duplicate metadata; adapter transfer, destination save/reopen,
and source metadata preservation. Thus the shared-part preservation fix now has
three-run desktop evidence on reported Office 16.0 x64. None failed at workbook
close in this run. All three stop at corruption fixture setup: LoadXML on an
already-loaded part returns 0xE0041808. Changed the fixture to delete its old part
through CustomerData and load deliberately unsupported XML into a fresh part.
Recovery and final second-edit assertions still require execution.

Ribbon investigation: no .vsto deployment package, CurrentUser code-signing cert,
or populated PowerPointIntelligence registration was found. The adapter runner
never installs or loads the VSTO add-in. Added scripts/Install-Local.ps1 for an
interactive desktop: requires PowerPoint closed normally, creates/reuses a
90-day CurrentUser development certificate (or accepts an explicit thumbprint),
builds signed manifests and runs VSTOInstaller with its normal trust prompt.
No root/publisher trust or Application Control changes are scripted. This is
local development installation, not production signing/distribution approval.
The script is syntax-checked but installation has not executed in this restricted
session. The new corruption fixture compiles; 27 local assertions pass.

## Installation wait investigation

Owner reports installer script stopped after assembly build message. Deployment
and application manifests now exist with two Signature elements each. This is
package-generation evidence, not installer success or cryptographic verification.
Process CIM access is denied and desktop processes are not visible to this
execution context, so the waiting stage cannot be established remotely.
Corrected interactive VSTOInstaller launch from Hidden to Normal, added explicit
progress/process ID, and added -SkipBuild to resume the already-generated package.
Installer performs normal trust/integrity checks. Script syntax passes; visible
installation and runtime loading remain pending. No existing desktop process was
terminated from this session.

## Desktop run: validation-20261022-121254-264d3533

All noninteractive checks pass. Office runs 2 and 3 pass every adapter assertion,
including malformed/missing metadata preservation, deleted-object recovery and
second edit/save/reopen. Run 1 passed through deleted-object recovery then failed
with 0x800706BA at Presentation.Slides during final reopen. Cleanup also failed
on the disconnected presentation. Result is TWO of THREE full runtime passes,
not a stable all-pass suite. No root cause for PowerPoint disconnection is proven.
Application crash-event log access from this session is denied.

Adjusted fixture cleanup so a COM cleanup error is recorded by HRESULT and does
not mask an existing failure; a cleanup error following success still fails the
run. Local 27 assertions pass and runner compiles. This does not fix or suppress
the original RPC failure. User reported VSTOInstaller success previously; ribbon,
task-pane operation and full process restart remain unverified.

## S0 stabilization follow-up

Inspection found unguarded managed calls in Activated/Closed/BeforeSave callbacks,
including pane removal while Office closes/switches documents. Added an event
boundary to contain exceptions and report only type/HRESULT to Trace. Save
cancellation is not changed. Tests verify successful execution and COM exception
containment without workbook text. This is a concrete hardening change, not
proof that it caused or fixes the observed RPC server disconnection.

Installer now increments the existing deployment revision for rebuilt payloads
so updated source is not confused with the prior installed ClickOnce version.
-SkipBuild intentionally reinstalls the existing payload and must NOT be used
to validate these new source changes. Validation runner captures sanitized Office
crash event metadata (time, application, module, exception code) when accessible;
no event messages or document contents are retained. Missing event access does
not silently become runtime success. Script syntax checks pass.

Runtime acceptance remains open: update installation with PowerPoint closed via
Install-Local.ps1, then run Test-All.ps1 -InteractiveOffice in the desktop session.
The session's PowerPoint activation restriction still prevents autonomous runtime
execution here. Full process restart and ribbon/task-pane checks remain separate
required S0 acceptance checks. No further S1 implementation was performed.
