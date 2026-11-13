# OPS-02 — COM automation interface

**Evidence category:** source verified; execution status is stated separately. Baseline includes the actual inspected checkout, not an assumption that every working byte equals HEAD.

## Design

| Dimension | Implementation-grounded description |
|---|---|
| 1. Purpose | Expose bounded diagnostic and chart insertion operations to local automation through COMAddIns.Object. |
| 2. Implementation status | Implemented local COM surface; runtime call acceptance requires Office execution. |
| 3. Internal components/classes/interfaces | IDemoAutomation IDispatch interface; DemoAutomation class with ClassInterface.None, stored PowerPoint application/creating-thread identity and read-only delegates. |
| 4. Important functions/methods | GetVersion, GetBuildInfo, GetRibbonXml, GetRibbonLoaded, GetContextPropertiesState and InsertChartJson. |
| 5. Inputs and outputs | JSON S0 ChartDocument for insertion; outputs new GUID or immutable strings/bool observations. |
| 6. Dependencies | PowerPoint PIA, DataContractJsonSerializer, MetadataCodec, PowerPointHost, DemoRibbon and DemoTools; ThisAddIn.RequestComAddInAutomationService creates the instance. |
| 7. Internal workflow | Read methods report source/build/UI snapshot; Insert checks same managed thread plus STA, bounds input, deserializes/validates S0, allocates new ID/resetrevision, inserts native chart and returns ID. |
| 8. External integrations | Local IDispatch COM only. It is not a REST/AI RPC protocol or cross-tenant service. |
| 9. Exceptions and error handling | Wrong thread/apartment throws InvalidOperationException; bad input size/model throws before insertion; COM/host failures propagate. No interface retry/timeout declared. |
| 10. Configuration | No separate API server setting; behavior follows loaded add-in build and active Office host/document. |
| 11. Security | No arbitrary code/member invocation. Snapshot read avoids Office/WinForms cross-thread calls. Windows/Office access supplies surrounding trust; no business RBAC. |
| 12. Tests | InstalledRibbon checks build/ribbon object; ContextPropertiesOfficeTests checks published state; host tests cover chart codec/insertion. Desktop-only portions not executed by this leaf. |
| 13. Known limitations | GetVersion says demo-v1/S0 JSON, distinct from canonical v2/v3; GetBuildInfo exposes local assembly path. No general action plan executor. |
| 14. Technical debt | Version/discovery/error envelopes are minimal; source-specific public diagnostics can reveal paths. |
| 15. Extension points | Add versioned typed operations only after host preflight/STA/selection checks and negative COM tests; never reinterpret v2/v3 models as S0 silently. |

## 16. Complete owned source map

| Source | Analysis | Responsibility/declarations |
|---|---|---|
| `apps/powerpoint-vsto/DemoApi.cs:1` | Complete text read; no runtime execution | Declarations: GetVersion, GetBuildInfo |
| `apps/powerpoint-vsto/Contracts.cs:1` | Complete text read; no runtime execution | Script/project entry; its actual flow and responsibilities are documented above. |

## 17. Handover and maintenance checks

Review the exact source/hash/line counts in `09-security/reviewed-files.csv` and this directory’s `source-module-mapping.csv`. Do not substitute historical PASS logs for current execution. Preserve unrelated module behavior and tests before modifying a boundary; new methods must have typed inputs, explicit error semantics and a corresponding negative test. See the dedicated AI/API/security/deployment documents for this module’s external contract.
