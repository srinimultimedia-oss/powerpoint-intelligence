# Application and tooling entry points

| Entry | Method/route | Source | Boundary |
|---|---|---|---|
| VSTO host generation | ThisAddIn.Designer generated host item | apps/powerpoint-vsto/ThisAddIn.Designer.cs:1 | Office loads signed payload; not a console Main |
| Startup and shutdown | OnStartup/OnShutdown | apps/powerpoint-vsto/ThisAddIn.cs:37 | Registers UI/host events and disposes owned resources |
| Ribbon extension | CreateRibbonExtensibilityObject/IntelligenceRibbon | apps/powerpoint-vsto/ThisAddIn.cs:326 | XML callback entry, returns DemoRibbon.Create |
| COM automation | RequestComAddInAutomationService/IDemoAutomation | apps/powerpoint-vsto/DemoApi.cs:12 | Local IDispatch object, thread-checked mutation |
| Standalone CLI | renderChartV2 main-module guard | packages/chart-engine/tools/render-chart-v2.mjs:1 | Reads v2 JSON; writes geometry-only scene to explicit output |
| Installer | Bootstrapper entry and Build-EasySetup | easySetup/Bootstrapper.cs:1 | Normal VSTO publisher flow, dualarchitecture selfcontained EXE |
| Local developer install | Install-Local.ps1 | scripts/Install-Local.ps1:1 | Signs/builds/installs; intentionally not invoked in docs run |
| Test orchestration | Test-All.ps1 | scripts/Test-All.ps1:1 | Snapshot/cached source proof, local checks and optinOffice |
| Library refresh | LibraryRefreshMonitor | src/Library.Core/LibraryRefreshMonitor.cs:1 | Local file watch/timed coordinator; not an enterprise scheduler |

No server Main, controller route, listener or deployed Azure gateway executable is present. Node and Console test Main methods are validation harnesses, not product network interfaces. The [symbol inventory](symbol-inventory.csv) includes declaration candidates for smaller internal/test modules and marks the lexical extraction limitation.
