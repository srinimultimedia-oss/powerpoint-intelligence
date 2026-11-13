# Debugging guide

| Symptom | First evidence | Responsible boundary | Safe next step |
|---|---|---|---|
| Ribbon absent | installed COM add-in status, payload marker/hash/receipt and VSTO prompt | VSTO registration and add-in load | Verify current signed package/architecture; inspect loader error before reinstalling |
| Commands disabled | compatibility status and deployed DLL config | ConfigureCompatibility | Correct known profile/build mismatch; do not infer product edition from 16.0 |
| Missing/ambiguous chart metadata | namespace count, fixed operation/HRESULT diagnostics | PowerPointHost, XmlPartUpdates | Preserve native objects; inspect attachment/tag association |
| Chart data not committed | linked diagnostic operation stages, point counts and verified source binding | LinkedChartData | Reproduce on desktop with one owned test deck; do not accept close alone as commit |
| Property box stale/hidden | immutable state snapshot: kind, foreground, selection identity, UI-STA | ContextPropertiesController | Bring intended PowerPoint window forward; distinguish focus prerequisite from mutation failure |
| Library search empty | configured sources, PIIDX validity, refresh error and path policy | LibraryRuntimeProfile/PiidxCatalog | Confirm deliberate source registration; do not fall back to missing proprietary caches |
| Renderer read fails after edit | native-binding verification versus canonical/scene hashes | PowerPointSceneSurface | Treat detected divergence explicitly; preserve the user's native edits |
| Cache runner missing | receipt, current hashes and scanner scope | recovery/build-integrity helpers | Review actual matching scanner paths/hashes before any source recovery |

Fixed-operation diagnostics intentionally avoid workbook contents. Catch boundaries should not be used as proof a whole mutation was atomic. `StartNewUndoEntry` groups native operations; it is not a transaction covering all COM calls. Use narrow source-specific integration tests. `apps/powerpoint-vsto/PowerPointHost.cs:34`; `apps/powerpoint-vsto/ThisAddIn.cs:133`; `apps/powerpoint-vsto/PowerPointHost.cs:14`; `src/Chart.Renderers/PowerPointSceneSurface.cs:14`.

COM failures may arise from the execution token (AppContainer), host focus, bitness, type-library registration, native workbook lifetime or busy Office calls. HRESULT alone is insufficient diagnosis. Never respond by disabling security tools, changing COM ACLs or force-terminating user Office processes. Historical activation evidence is environment-specific; rerun diagnostics only in an authorized developer session.
