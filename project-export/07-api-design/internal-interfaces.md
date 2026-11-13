# Internal interfaces

## COM automation surface

`apps/powerpoint-vsto/DemoApi.cs:12` declares six methods; interface GUID is a stable local COM identity and class uses ClassInterface.None. `ThisAddIn.RequestComAddInAutomationService` supplies this object. `apps/powerpoint-vsto/ThisAddIn.cs:29`.

| Method | Parameters / return | Validation / semantics |
|---|---|---|
| GetVersion | none → string | Literal demo-v1 / S0 ChartDocument JSON; not a semantic production API version negotiation |
| GetBuildInfo | none → string | Build marker, catalog length, module version ID and assembly path; diagnostics may reveal local path |
| GetRibbonXml | none → string | Current ribbon definition generated from source catalog |
| GetRibbonLoaded | none → bool | Published Office OnLoad observation; does not prove visual layout or every command |
| InsertChartJson | JSON string → new chart GUID string | Same creating UI thread and STA; size bounded by MetadataCodec; DataContract JSON deserialize; strict model validate; overrides identity and revision; invokes native host insertion |
| GetContextPropertiesState | none → XML string | Immutable UI-published snapshot delegate; deliberately no Office/WinForms operation on caller thread |

No retries/timeouts are declared on this COM interface. Thread misuse throws InvalidOperationException; blank/oversized or invalid models throw before host mutation. Office failures propagate from the host. It does not accept an arbitrary COM method name or executable code. `apps/powerpoint-vsto/DemoApi.cs:35`.

## Host interface

`IPowerPointHost` methods: GetSelectedSlideId() → string; GetSelectedObjectIds() → IReadOnlyList<string>; InsertEditableChart(ChartDocument) → IReadOnlyList<string>; UpdateEditableChart(chartId, document) → void; ReadChartMetadata(chartId) → ChartDocument; SaveChartMetadata(document) → void. `apps/powerpoint-vsto/Contracts.cs:7`. This is the S0 version-1 native clustered-column model, distinct from version-2/3 canonical domain contracts. Metadata is validated by MetadataCodec and persisted on native shapes. UI/host context and workbook operations are not interchangeable with DTO validation.

## Renderer surface

`IEditableSceneSurface`: VerifyAccess; Begin(operationId,chartId); CreateAnchor(ScenePlot,bool) → int; CreateNode(SceneNode,token) → int; Group(IReadOnlyList<int>) → int; Persist(rootId,chartId,xml); Rollback() → bool; Complete(). `src/Chart.Renderers/NativeRenderingCore.cs:136`.

NativeSceneRenderer validates canonical/scene correspondence before Begin, rejects reentrancy, tracks the active stage and wraps failures in NativeRenderException with rollback completeness. A false rollback result is observable; it is not silently counted as successful cleanup. Surface owns COM RCWs; coordinator passes validated DTOs/IDs.

## Event and IPC contracts

ThisAddIn subscribes WindowSelectionChange, WindowBeforeDoubleClick, PresentationOpen, PresentationBeforeSave, WindowActivate and PresentationClose, and unsubscribes during shutdown. Layout and agenda timers coalesce some edits. `apps/powerpoint-vsto/ThisAddIn.cs:44` `apps/powerpoint-vsto/ThisAddIn.cs:77`.

`OfficeMutationLease.TryEnter` serializes native automatic writes across processes using a named lease derived from the Office PID. It is concurrency coordination, not user authentication. Test runners use the same mutation lease; contextual read observation remains separate. PowerPoint selection snapshots reject stale identity before applying a queued property edit; see PowerPoint integration docs.
