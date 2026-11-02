# v17 parity increment: implementation and validation

Status: **partial implementation; full think-cell parity is not accepted**.

## Executed local evidence

Source marker: `PPI-DEMO-RIBBON-20261026-v17`.
Source SHA256: `DF745CA44CDA8C2B69347E24860FD782028A19F41DC4DB84546D0879CEC4C8D5`.
Run: `artifacts/validation-20261027-000225-a2768410/results.json`.
Command: `scripts/Test-All.ps1` without InteractiveOffice.

**20 PASS / 1 FAIL**, all AnyCPU/x86/x64 compile targets passing. This supersedes the previous v16 local result of 19 PASS / 2 FAIL for this implementation increment only. The ribbon-definition failure was a .NET Framework PathTooLongException caused by the test fixture directory and temporary filename under the nested source snapshot. Shorter fixture/temp prefixes fix it; no global Windows path/security policy was changed.

Within passing checks:

| Check | Executed result |
|---|---|
| Native renderer core | 41 assertions pass; geometry/mutation abstraction, not Office COM |
| Renderer ribbon preparation | 9 pass |
| Managed agenda planner | 654 pass |
| Chart annotation geometry/metadata | 34 pass |
| Formatting validation/cancellation/rollback | 34 pass |
| Quality-review rules/bounds | 31 pass |
| Native axis/series validation | 27 pass |
| Explicit Excel mapping/codec/shape-change policy | 32 pass |
| Data source/reference contracts | 78 pass |
| Link inventory/association policy | 28 pass |
| Service profiles | 41 pass; no authenticated/network request |
| Ribbon bindings/icons | 90 existing tool bindings plus 26 workflow bindings; 140 button/menu icon render checks |

The sole local FAIL is **scene-contracts-x64**, stopped at the provenance guard. The stable executable was retained rather than overwritten or copied to another path. Its inspected SHA256 remains `EA5C34B16761FD30AC9ADF0F8CB4DE362B1AA81A93DD9D170D06EA0A087FA7A3`, matching the previously recorded McAfee quarantine hash. A reviewed provenance/detection resolution is required; this report does not infer a new current detection or a false positive. No AV setting, exclusion or receipt was changed to obtain a pass.

## New implementation

- Native section agenda insertion/explicit refresh with stable hyperlinks and editable fields.
- Native chart-bound reference/difference annotations with explicit refresh; supported primary unstacked linear axes only.
- Linked text-capable native tables, bounded to 400 cells.
- Primary value-axis bounds/interval/number format/reversal/logarithmic mode and selected-chart scale matching.
- Bounded native column/line/area series conversion and secondary-axis selection, refusing incompatible managed S0 chart contracts.
- Native selected text/group/table formatting and RGB color pickers with preflight, Undo boundary and bounded restoration plan.
- Read-only quality review and Locate with explicit limits; no cleanup/sanitization claim.
- Native PDF export of the current presentation; no messaging/sharing side effects.
- Explicit Excel worksheet/range/column mapping in v2 metadata, legacy v1 decode, manual compatible chart refresh, confirmed selected-chart relinking and a real source/object inventory.
- Nonsecret service settings UI; external authenticated features are still unimplemented/unconfigured.

AD-017 records the owner's explicit mapping/relink policy and live-service scope. The main parity audit records the remaining capabilities. No vendor binaries/protected internals were inspected, and no vendor artwork was added to the product.

## Desktop regression scenarios implemented, not executed here

`WorkflowOfficeTests.cs` checks owned-deck agenda targets/edit/refresh/persistence, native reference/difference calculation and refresh, duplicate identity refusal, native formatting/axis persistence, series conversion, read-only quality object preservation and native PDF export. Hyperlink targets are checked structurally; actual slideshow click-through remains a manual check.

`MappedExcelOfficeTests.cs` checks an actual non-A1 workbook range, column projection, source edits/manual chart refresh, one association after repeat refresh, header-change refusal before chart mutation and source/native-value save/reopen. Relink and picker/dialog interaction remain manual acceptance scenarios.

Both compile into the existing `DemoOfficeTests.exe`; no standalone alternate executable was created. The stable scene/template runner paths remain unchanged. Compilation does not establish the native results above.

This agent's confirmed AppContainer cannot activate PowerPoint. Desktop-token activation previously passed. That evidence supports the execution-boundary diagnosis, not a claim that the new installed ribbon or Office scenarios already work. The owner agreed to run the frozen-source installation and interactive suite in normal Windows PowerShell with McAfee active:

```powershell
.\scripts\Install-Local.ps1
.\scripts\Test-All.ps1 -InteractiveOffice
```

Save/close user presentations before testing. Complete the visible publisher prompt. The suite checks the loaded v17 marker/source hash, all workflow groups and bindings; retain its artifacts folder. Source is frozen while this run is pending.

## Remaining acceptance gates and feature gaps

- New installed-ribbon interaction, dialogs, editable native output, rollback/failure behavior and save/reopen need genuine desktop execution. Office 2016, Office 2019 and Microsoft 365 runtime certification remains separate from architecture compilation.
- x64 scene provenance/McAfee review remains unresolved.
- Prior original-template pixel-fidelity/save-reopen and demo-runtime failures remain open until a corresponding fresh native run resolves them.
- Complete advanced chart interaction, profile orientation, waterfall totals/connectors, calendar Gantt, Mekko editing, axis breaks, semantic label collision/drag/layout, automatic annotations, hierarchical/synchronized agendas, general constraint layout, shared slide workbook, complete link lifecycle, cleanup/sanitized sharing, batch automation and authenticated online/AI/provider services remain incomplete.
- The vendor installed-version feature inventory and public manuals are bounded evidence; the audit is not exhaustive certification of every think-cell version or edition.

References: public [axis behavior](https://www.thinkcell.com/en/resources/manual/axes), [native axis bounds](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.axis.minimumscale), and [native PDF export](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat). Public documentation establishes desired/API behavior; it does not substitute for execution evidence.
