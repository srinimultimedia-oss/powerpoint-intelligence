# API and integration inventory

No implemented HTTP server or REST endpoint exists in the inspected source. `services/ai-gateway` is a schema/README scaffold; there are no controllers/routes/ports/listeners. Therefore no fictional OpenAPI service specification is emitted. See [OpenAPI applicability](openapi/README.md).

| ID | Interface | Direction / consumer | Implementation / status |
|---|---|---|---|
| API-01 | IDemoAutomation (IDispatch COM) | PowerPoint COMAddIns.Object → DemoAutomation | `apps/powerpoint-vsto/DemoApi.cs:12`; implemented, local automation |
| API-02 | IPowerPointHost | Add-in commands → PowerPointHost | `apps/powerpoint-vsto/Contracts.cs:7`; managed internal boundary |
| API-03 | IEditableSceneSurface | NativeSceneRenderer → PowerPointSceneSurface or test fake | `src/Chart.Renderers/NativeRenderingCore.cs:136`; managed adapter |
| API-04 | IRibbonExtensibility and callback methods | Office → IntelligenceRibbon → ThisAddIn | `apps/powerpoint-vsto/ThisAddIn.cs:332`; Office callback contract |
| API-05 | PowerPoint/Excel COM object models | In-process add-in → Office native document/workbook | `apps/powerpoint-vsto/DemoLinks.cs:50`; desktop only |
| API-06 | Win32 window/token and SQLite native P/Invoke | Controllers/diagnostics/catalog → Windows | NativeSqlite, ContextPropertiesView/Controller, diagnostic C# snippet; OS-native interfaces |
| API-07 | Linked data files | User-selected XLSX/CSV/JSON/XML → LinkedTable/mapped chart | DataSources, ExcelRangeMapping, MappedLinkOperations; file contracts, not REST |
| API-08 | PIIDX catalog and configuration | Local gallery → SQLite/files/watchers | Library.Core; independent local persistence, see data design |
| API-09 | Versioned CustomXMLPart metadata | Add-in ↔ presentation/shape CustomerData | Host/domain/rendered/text/agenda/indicator/link codecs; persistent file interfaces |
| API-10 | Signed VSTO deployment | Setup/Install-Local → VSTOInstaller | `easySetup/Bootstrapper.cs:33`; visible trust prompt, per-user |
| API-11 | Service connection profiles | User UI ↔ local XML | `apps/powerpoint-vsto/ServiceConnections.cs:89`; no network integration |
| API-12 | AIActionPlan | Planned gateway/action orchestrator | JSON Schema only; not an executing interface |

No independent API authentication layer exists for COM calls. Windows/Office process access and add-in loading establish the surrounding trust; mutations have thread, model and selection guards, not RBAC.

See [internal interfaces](internal-interfaces.md), [external interfaces](external-interfaces.md), [contracts](request-response-contracts.md). This inventory counts integration boundaries, not every private helper method; the repository-wide symbol inventory provides exhaustive declarations.
