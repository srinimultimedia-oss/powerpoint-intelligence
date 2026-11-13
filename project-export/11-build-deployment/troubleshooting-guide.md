# Troubleshooting

| Symptom | Evidence and source-grounded response |
|---|---|
| Missing ribbon | Inspect COM Add-ins Connect/load error, build marker/assembly hash, GetRibbonLoaded/OnLoad and install receipt. Registration/load/visual runtime are distinct. |
| COM0x80070005 | Collect token/registration/stage with Diagnose-PowerPointActivation; compare normal desktop token. Historical AppContainer denials do not prove antivirus is cause. Do not weaken controls speculatively. |
| Type library/bitness errors | Check Office API/executable build/process architecture and normal Office/VSTO registration. PIAs version15 are compile API identities, not Office 15 support. |
| Workbook RPC rejection/disconnection | Fixed step/HRESULT, owned crash metadata and LinkedChartDiagnostic; respect bounded retries/lifetime/mutation lease. Timeout is not successful commit. |
| Property panel hidden/stale | Published state reason/stage/foreground/selection identity; desktop foreground prerequisite and stale-target guard prevent wrong mutation. |
| STALE PACKAGE | Compare current marker/manifest/DLL hash/deployment version. Force current signed build; SkipBuild does not fix stale payload. |
| Missing certificate/UnknownError | Check approved normal desktop CurrentUser/My key/EKU/trust. AppContainer cert views differ; failed builder preserves old EXE. |
| Payload checksum after atomic EXE replacement | Verify receipt/EXE and read current immutable bytes. Resource reader avoids Assembly.LoadFile path identity caching; regression exists. |
| Watcher idle after publication | Normal polling; Ctrl+C stops. Failure waits for a new source edit and keeps prior EXE. |
| Missing scene/template/PIIDX runner | Preserve receipt/hash/path/launch status; review scanner scope and exact action before explicit source recovery. Do not delete receipt to force rebuild. Current collector scope excludes PIIDX (SEC-02). |
| Moved/changed data | Header/range/path/mapping check; stop refresh, reselect and confirm. Never assume automatic source rebinding/sample fallback. |
| Library unavailable | Check explicit root/index/config and permissions/integrity/FTS; no silent fallback from an unavailable configured root. |
| Fresh package looks stale in PS7 | Compare NLS/ICU sorting host; use documented normal Windows PowerShell, not bypassing freshness. |

Evidence: `scripts/Test-InstalledRibbon.ps1:1` `scripts/Diagnose-PowerPointActivation.ps1:1` `apps/powerpoint-vsto/PowerPointHost.cs:214` `apps/powerpoint-vsto/DataSources.cs:197` `easySetup/Common.ps1:2`.

Diagnostic scripts can write artifacts/temp or activate Office; they were analyzed, not executed by this export. Keep incident output to bounded identity/error metadata; do not publish workbook contents/private credentials.
