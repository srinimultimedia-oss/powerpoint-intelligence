# Office regression repair — elements2

Source identity: `PPI-DEMO-RIBBON-20261028-v17-elements2`. Native acceptance is pending a new desktop run; the fixes below are not claimed as executed native passes.

## Desktop evidence inspected

Owner run: `artifacts/validation-20261028-103116-c3ea917e`, source snapshot `1E77B7314AB9379CCF3E89F0529553D026334930105130D5A0B41FD0062DD5B4`.

- Installed ribbon, native shapes and all three S0 Office runs passed.
- Indicators passed **16 native assertions**, including editable vectors, duplicate-state independence, native child formatting, save/reopen and further edits. This is acceptance evidence for those tested scenarios on this Office16.0/x64 machine, not other Office versions or untested shortcuts/dialogs.
- Text, grid and connectors failed on reloading a populated CustomXMLPart. Agenda used the same invalid update/rollback pattern, in addition to the observed frame-size rejection.
- Workbook close threw `0x80028018` through reflection dispatch with the invariant culture. Subsequent demo runs lost the Office RPC server; most of the 90 commands were blocked rather than tested.
- Symbol fixture attempted to leave text mode using Slide.Select; the observed selection remained a text range. The test therefore exercised replacement again, not new-box insertion.
- Workflow series-value verification hit a typed/dynamic SAFEARRAY cast error.
- Template insertion/reopen failed exact raster comparison on Pyramids source slides4 and7. All source/insertion native geometry comparisons passed.
- x64 scene contracts stopped at the existing provenance guard. This log does not show a new antivirus detection.

## Source corrections

`XmlPartUpdates` stages a new XML part, loads and verifies it, then removes the prior association. Shape updates capture other owners and restore independently if detaching a shared part removes their metadata. Registry updates create the replacement before deleting the prior registry. The caller's rollback resolves the current part rather than invoking LoadXML on an already-loaded/stale part. Text, grid, connector and chapter update/restore paths now use this helper. New-part initialization still uses LoadXML. XML remains bounded and DTDs/external entities are prohibited. The [Microsoft LoadXML example](https://learn.microsoft.com/en-us/office/vba/api/office.customxmlpart.loadxml) initializes a newly added part; it is not an overwrite API.

The agenda's empty layout frame explicitly disables both native AutoSize modes and reapplies its intended bounds. A rejected frame now reports its coordinates and slide dimensions. The prior log did not record actual frame bounds, so the precise sizing trigger is still unverified; the patch prevents empty-text autofit from sizing the frame and adds evidence for any remaining failure.

Workbook reflection close now uses en-US/LCID1033 only for that invocation, without changing thread, user or Office settings. The [Microsoft Excel globalization guidance](https://learn.microsoft.com/en-us/visualstudio/vsto/globalization-and-localization-of-excel-solutions) documents LCID1033 and explicit InvokeMember culture. The exact desktop failure is consistent with the invariant-locale dispatch mismatch; successful native closure still needs execution evidence. The code does not treat TYPE_E_INVDATAREAD or lost RPC as a successful close.

The symbol fixture explicitly calls Selection.Unselect and asserts it has left text mode. Workflow validation explicitly types the Series and uses the existing SAFEARRAY normalizer. A failed RPC call (`0x800706BE`) now blocks subsequent dependent native scenarios immediately; blocked rows are not counted as passes.

## Template raster diagnosis — strict status retained

Offline comparison of the saved native PNG evidence found:

| Fixture | Insert/source changed pixels | Maximum channel delta | Unchanged source re-export changed pixels | Reopen/source changed pixels | Reopen maximum channel delta |
|---|---:|---:|---:|---:|---:|
| original6 (Pyramids slide4) | 2 | 6/255 | 2 | 8 | 10/255 |
| original9 (Pyramids slide7) | 2 | 6/255 | 2 | 8 | 10/255 |

Differences occupy tiny gray icon-edge regions; insertion and unchanged-source repeated export have the same two-pixel differences from the first source image. This establishes source-render nondeterminism at this resolution. It does not justify silently weakening the owner's exact-fidelity requirement. Pixel acceptance is unchanged; bounded antialias tolerance was explicitly offered to the owner and remains pending. Source template files and template runner/cache receipts were not modified.

## Validation / next native run

Eight additional pure assertions cover safe XML inputs and Excel LCID isolation under an ar-SA caller culture. Existing element contracts and native fixtures are retained. New native fixture assertions compile but have not executed in this AppContainer session. Stable DemoContractTests.exe and DemoOfficeTests.exe paths are reused; no new runner EXE or security exception is introduced.

With personal presentations saved/closed and antivirus active, use normal desktop PowerShell:

```powershell
.\scripts\Install-Local.ps1
.\scripts\Test-DemoFeatures.ps1 -InteractiveOffice -ElementsOnly
.\scripts\Test-DemoFeatures.ps1 -InteractiveOffice -AgendaOnly
.\scripts\Test-All.ps1 -InteractiveOffice
```

The first two focused native suites give short, isolated feedback before the full regression. Full acceptance remains open for the x64 scene provenance hold, strict template raster fidelity, repaired native scenarios, full feature gaps and other Office runtime profiles.

## Final local validation of repair

Run: artifacts/validation-20261028-104331-9b0f70e5/results.json. Source SHA256: `4A73B647C813452EFD3E213673DF14B97D3B4D5122C103ABEE0B2C1E7078A512`.

Result: **20 PASS / 1 FAIL**. Only scene-contracts-x64 remains held by the unchanged provenance policy. The 864 element/productivity assertions plus 8 XML/Excel-locale assertions pass (**872** total); renderer core remains 41/41. All three .NET Framework architecture targets compile, and each actual snapshot DLL was separately checked for the elements2 marker. The native runner compiles. No repaired native scenario was executed in the agent session. Final source fingerprint matches this snapshot and git diff --check passes.

Next install must rebuild the current elements2 source; the prior installed-ribbon PASS belongs to the owner's earlier snapshot. Strict template pixel status and x64 scene provenance status remain unresolved. No security controls, scanner settings, exclusions, source-template assets or runner receipts were changed.

### Owner-approved raster policy follow-up

Owner approved bounded antialias tolerance on 28 October 2026 (AD-018). The strict raster-only condition above is superseded by [the bounded fidelity policy](template-fidelity-policy.md), which retains exact native checks and separate exact/tolerant metrics. Saved raster evidence replays pass; new native Office acceptance remains pending.
