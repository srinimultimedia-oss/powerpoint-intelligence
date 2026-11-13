from pathlib import Path
import json,csv,re,datetime,hashlib,subprocess,os
R=Path(__file__).resolve().parents[2];E=R/'project-export'
def ref(path,needle=None):
 text=(R/path).read_text(encoding='utf-8-sig');ls=text.splitlines();i=next((i for i,s in enumerate(ls,1) if needle is None or needle in s),None)
 if i is None:raise ValueError((path,needle))
 return f'`{path}:{i}`'
def put(path,title,text):
 (E/path).write_text('# '+title+'\n\n'+text.strip()+'\n',encoding='utf-8')
startup=ref('apps/powerpoint-vsto/ThisAddIn.cs','private void OnStartup');shutdown=ref('apps/powerpoint-vsto/ThisAddIn.cs','private void OnShutdown');run=ref('apps/powerpoint-vsto/ThisAddIn.cs','private void Run(Action');host=ref('apps/powerpoint-vsto/PowerPointHost.cs','public sealed class PowerPointHost');lease=ref('apps/powerpoint-vsto/PowerPointHost.cs','internal sealed class OfficeMutationLease');event=ref('apps/powerpoint-vsto/PowerPointHost.cs','internal static class OfficeEventBoundary');api=ref('apps/powerpoint-vsto/DemoApi.cs','public interface IDemoAutomation');renderer=ref('src/Chart.Renderers/NativeRenderingCore.cs','public sealed class NativeSceneRenderer');surface=ref('src/Chart.Renderers/PowerPointSceneSurface.cs','public sealed class PowerPointSceneSurface');domain=ref('src/Chart.Domain/Chart.Domain.csproj','<Compile');schema=ref('src/Library.Core/PiidxSchema.cs');catalog=ref('src/Library.Core/PiidxCatalog.cs');source=ref('src/Library.Core/PresentationSource.cs');services=ref('services/ai-gateway/README.md');settings=ref('apps/powerpoint-vsto/App.config');decisions=ref('docs/DECISIONS.md')
context='''```mermaid
flowchart LR
  U[Presentation author] --> PPT[Windows PowerPoint desktop]
  PPT --> NW[NexWave VSTO add-in]
  NW --> DOC[Current presentation native objects and custom XML]
  NW --> FS[User-selected Excel CSV JSON XML files]
  NW --> LIB[Local PPTX POTX template libraries]
  NW --> IDX[Per-user PIIDX SQLite catalog and preview files]
  NW --> CFG[Local connection profile configuration]
  CFG -. preparation only .-> EXT[Azure AI Tableau MCP and asset providers]
```
'''
put('02-architecture/enterprise-architecture.md','Enterprise architecture',f'''
**VERIFIED — source architecture, not enterprise deployment certification.** The running product is a Windows desktop add-in hosted in the interactive PowerPoint process. Authors invoke ribbon callbacks, dialogs, a task pane and selection-following properties. Its primary data boundary is the active presentation, with user-selected files and a per-user local catalog as supporting stores. The Microsoft Office process, its COM apartment, VSTO loader, .NET Framework and Windows filesystem are platform dependencies. {startup}; {host}.

{context}
The context diagram's dotted provider edge is configuration preparation, not a live inference or integration path. `services/ai-gateway` contains a schema and README; no deployable web service, database migration service, tenant middleware or identity server exists there. {services}. Treat on-prem hosting and Azure OpenAI as approved intentions whose implementation and network/data-classification decisions remain open. {decisions}.

Trust boundaries are (1) signed installation into an Office user session; (2) file parsing and canonical model validation before mutation; (3) COM UI/STA ownership; (4) presentation metadata and native-object reconciliation; and (5) package/index integrity checks for local libraries. The cross-process mutation mutex coordinates cooperating add-in/test writers, not malicious processes or Office's own internal operations. {lease}.

No RBAC/SSO/tenant isolation or encrypted enterprise storage can be inferred from local profile configuration. Sharing generated reports, licensed assets and source workbooks remains an organizational process. See [security](../09-security/security-architecture.md), [AI](../06-ai-architecture/ai-solution-architecture.md) and [deployment](../11-build-deployment/deployment-guide.md).''')
put('02-architecture/solution-architecture.md','Solution and container architecture',f'''
The unit of desktop deployment is the signed VSTO manifest/payload, loaded by PowerPoint. C# source is compiled from explicit project items, including linked domain/rendering/library sources. There is no `.sln` in this baseline; the two `.csproj` files are the build entry points. A source folder's presence does not establish inclusion: V3 contracts are absent from both active C# project compile lists. {domain}; {ref('apps/powerpoint-vsto/PowerPointIntelligence.csproj','<Compile')}.

```mermaid
flowchart TB
  subgraph Desktop[Interactive Windows desktop]
    PPT[PowerPoint process] --> VSTO[VSTO lifecycle and ribbon]
    VSTO --> UI[WinForms task pane dialogs contextual controls]
    UI --> TOOLS[Command and editing modules]
    TOOLS --> HOST[PowerPoint COM adapter]
    TOOLS --> MODELS[Canonical codecs and validation]
    TOOLS --> LIB[Library catalog and native import]
    LIB --> SQL[Windows winsqlite3 SQLite]
    LIB --> CACHE[Local files and preview cache]
    HOST --> PPT
  end
  subgraph Dev[Build and validation process boundary]
    NODE[Node TypeScript calculations and scene contracts]
    PS[PowerShell build and test scripts]
    CSC[Visual Studio Roslyn and MSBuild]
    PS --> CSC
    PS --> NODE
  end
  NODE -. canonical JSON and scene JSON artifacts .-> MODELS
```

Node tooling validates deterministic calculations and emits a geometry-only scene when its CLI is invoked. The production VSTO adapter is not documented as launching Node on every chart command; C# supplies its own validated clustered projection. {ref('packages/chart-engine/tools/render-chart-v2.mjs','export async function')}; {ref('src/Chart.Renderers/NativeRenderingCore.cs','public static class ClusteredProjection')}.

State is split across PPTX custom XML/tags, linked local input files, local JSON/XML configuration, PIIDX SQLite and PNG previews, plus build/cache receipts. There is no deployed on-prem application database in the repository. Local PIIDX is implemented and is separate from the unresolved enterprise database choice. {schema}; {decisions}.''')
put('02-architecture/application-architecture.md','Application architecture',f'''
## Startup and lifecycle

VSTO creates the generated `ThisAddIn` host item and calls the application's startup handler. That handler constructs the host adapter, starts layout timing, prepares an agenda timer, evaluates Office compatibility, creates the context-property controller, subscribes selection/double-click/open/save/activate/close events and invalidates an existing ribbon. The 750ms layout and 650ms agenda intervals are WinForms UI scheduling, not backend workers. {startup}.

Compatibility loads the deployed DLL configuration and compares API major 16, requested architecture and optional executable build floor. An Office API string of 16.0 cannot distinguish 2016, 2019 and Microsoft 365 licensing/update channels. Invalid configuration disables document commands rather than guessing. {ref('apps/powerpoint-vsto/ThisAddIn.cs','private void ConfigureCompatibility')}; {ref('apps/powerpoint-vsto/OfficeCompatibility.cs')}.

## Command execution

Ribbon callback objects forward to `ThisAddIn`, which applies `busy`, host, compatibility and mutation-lease checks. `EnsurePane` creates/rebinds a task pane to the current window on demand. User action code executes while holding a named mutex for the PowerPoint PID. Specific failures become status text; a generic catch reports operation unavailability. Cancellation is a no-op. The `finally` path resets busy and requests a context refresh. {run}.

Selection changes request a property snapshot, invalidate controls and restart the debounce timer. Activation avoids creating/discarding panes or rewriting metadata directly in the COM callback; it marks pane refresh. Before-save work targets only the active presentation and refreshes managed agenda/layout/chart state. These handlers use an exception boundary to prevent managed exceptions escaping Office callbacks. {ref('apps/powerpoint-vsto/ThisAddIn.cs','private void BeforeSave')}; {event}.

## Shutdown and resource ownership

Shutdown disposes the context controller, timers and layout form, unsubscribes all registered application events and removes the task pane. Native COM object lifetime is handled locally in modules; the repository does not implement a general reference-count tracker proving every RCW release. Never release shared Office application objects indiscriminately. {shutdown}. See [interoperability](../04-powerpoint-integration/interoperability.md).

## Runtime limits

The floating property box is foreground/selection-sensitive and can hide while Excel owns focus. Mutation coordination does not make COM calls thread-safe. Live service operations, enterprise scheduling, authorization and model invocation are not implemented merely because command labels exist. Details are mapped in [feature status](../05-features/feature-status-matrix.md).''')
put('02-architecture/component-architecture.md','Component architecture',f'''
```mermaid
flowchart LR
  R[DemoRibbon and IntelligenceRibbon] --> A[ThisAddIn]
  A --> C[ContextPropertiesController and ActionRouter]
  A --> T[DemoTools ElementCommands ProductivityTools]
  A --> H[PowerPointHost]
  T --> D[DataSources ExcelRangeMapping MappedLinkOperations]
  T --> E[TextShapeElements Indicators Connectors ProcessFlow]
  T --> G[ChapterAgenda PresentationAgenda]
  A --> L[LocalLibraryWorkspace and Gallery]
  L --> P[PIIDX core]
  T --> N[Native chart controls annotations and formatting]
  A --> S[RendererRibbonCommand]
  S --> V[ClusteredProjection and NativeSceneRenderer]
  V --> F[IEditableSceneSurface]
  F --> O[PowerPointSceneSurface]
```

The graph shows principal call direction, not exhaustive dependency injection. Most features are static/helper classes receiving a PowerPoint Application/Shape/Selection; there is no general DI container, plugin discovery framework or durable command-bus service. The host interface and editable-surface interface provide narrower test seams. {ref('apps/powerpoint-vsto/Contracts.cs','public interface IPowerPointHost')}; {ref('src/Chart.Renderers/NativeRenderingCore.cs','public interface IEditableSceneSurface')}.

`Demo` names are historical implementation identifiers in active production compile items, not proof a source is unused. Conversely V3 model files and the AI schema are not evidence of a running product capability. COM automation exposure is explicitly limited to `IDemoAutomation`; it is a local add-in automation object, not a REST API. {api}.

See [module inventory](../03-modules/module-inventory.md) for every coherent module and [dependency matrix](../03-modules/module-dependency-matrix.md) for evidence and test seams. Package namespaces preserve `PowerPointIntelligence` for compatibility while visible branding is NexWave.''')
put('02-architecture/deployment-architecture.md','Deployment architecture',f'''
```mermaid
flowchart LR
  DEV[Approved Windows build workstation] --> CSC[Visual Studio MSBuild and Roslyn]
  KEY[Approved certificate private key] --> CSC
  CSC --> X86[Signed x86 VSTO payload]
  CSC --> X64[Signed x64 VSTO payload]
  X86 --> EXE[Self-contained signed setup EXE]
  X64 --> EXE
  EXE --> USER[Licensed target Windows user]
  USER --> INST[Per-user extraction and VSTO publisher prompt]
  INST --> OFFICE[Office process loads matching architecture]
  OFFICE --> DATA[Presentation and per-user PIIDX data]
```

The bootstrapper detects target Office architecture, validates/extracts its embedded bundle and invokes the VSTO installation path. Existing prerequisite .NET Framework 4.8/VSTO runtime and PowerPoint are required. It does not silently fetch or install prerequisites. The private key belongs to the build identity, never to the distributable. Exact mechanics and rollback paths are in {ref('easySetup/Bootstrapper.cs')} and {ref('easySetup/Build-EasySetup.ps1')}.

Local development installation uses `scripts/Install-Local.ps1`, a separate path from distributed easySetup extraction. Both require appropriate publisher trust. The exporter performs no installation and treats prior deployment receipts as historical evidence. A local-development signer trusted on one machine is not evidence of an enterprise publisher trusted elsewhere. {ref('scripts/Install-Local.ps1')}.

No cloud/on-prem API container is deployed by these scripts. Network identity, enterprise database and Azure topology remain open decisions. Keep installed extraction locations because VSTO may continue referring to them; rebuilding/removing artifacts indiscriminately can break local registrations. [Packaging details](../11-build-deployment/packaging-installation.md).''')
put('02-architecture/architecture-decisions.md','Implemented decisions and unresolved architecture',f'''
The controlling decision register is {decisions}. This export does not update or silently resolve it.

| Decision family | Observed implementation | Boundary |
|---|---|---|
| AD-001/006 Windows VSTO | Framework 4.8 C# add-in and Office COM adapters | Actual estate/version certification is not established by a profile string |
| AD-004/009 native editability/metadata | Native Office charts/shapes and CustomXMLPart plus tags | Metadata validity is checked; native edits can create detectable divergence |
| AD-008 deterministic TypeScript | Six basic engine families, strict scene contracts, offline CLI | Not a complete native renderer for all required chart families |
| AD-013 grouped shapes | Native scene surface and original template compositions | Grouped shapes are editable drawing objects, not native Excel-backed chart types |
| AD-014/017 explicit linked data | Local source selection and confirmed worksheet/range mapping | Changed headers/missing files stop refresh; no enterprise scheduling |
| AD-015/016/019 libraries | Procedural 520-layout catalog and local original-file source library | Different content modes; licensed source reuse is not independently authored parity |
| AD-018 raster tolerance | Bounded edge-local antialias comparison plus native geometry/format checks | Sixteen original slides are a sampled historical fidelity test, not all 497 |
| AD-020/021 setup and brand | Dual-bitness self-contained VSTO setup and NexWave display name | Compatibility IDs/storage names remain PowerPointIntelligence |
| AD-002/003/007 on-prem and AI | Service schema/profile scaffolds | Production service, Azure invocation, RBAC/SSO remain unimplemented |

Open Q-001..Q-012 include exact tested Office builds, identity, Azure resource/networking/data policies, enterprise infrastructure, signing distribution trust, licensing and performance/locale policy. Local SQLite implementation does not resolve enterprise Q-006. Configurable URLs do not resolve auth/data-governance Q-003..Q-005. Historical owner S0 acceptance is a project decision, not fresh runtime evidence in this export.

Design tension: active procedural/native/original-library paths and legacy S0/v2/V3 contracts coexist. Extension work must choose a production representation and codec intentionally, preserve source compatibility identities and add explicit migration/round-trip tests before claiming a new persisted version. {domain}; {ref('src/Chart.Domain/S0Migration.cs')}.''')
put('02-architecture/architecture-diagrams.md','Architecture diagram index',f'''
The architecture diagrams are maintained with their interpretation and source evidence:

| View | Document | Meaning |
|---|---|---|
| C4 context | [Enterprise](enterprise-architecture.md) | Author, PowerPoint, local sources/catalog and provider-preparation boundary |
| C4 containers/processes | [Solution](solution-architecture.md) | Office/VSTO process, local storage, development Node/compiler processes |
| C4 components | [Components](component-architecture.md) | Lifecycle/UI, feature handlers, codecs, PIIDX and editable surface |
| Deployment | [Deployment](deployment-architecture.md) | Dual-bitness signed payload build and per-user VSTO activation |
| Execution sequences | [Sequences](../10-workflows/sequence-diagrams.md) | Host events, linked writes, library operations and teardown |
| Activities/trust boundaries | [Flows](../10-workflows/activity-diagrams.md) | Validation, current-window mutation and failure branches |
| Data/ERD | [Data classes](../08-data-design/class-diagrams.md) | Actual model and PIIDX entities |

All graphs describe inspected code boundaries, not measured throughput, HA topology, unimplemented Azure invocation or third-party feature parity. Mermaid fences are inventoried and structurally checked. Full syntax parsing/rendering is separately reported; lack of a local Mermaid runtime must not be called a parser pass.''')
put('13-maintenance/developer-onboarding.md','Developer onboarding',f'''
1. Read export README, [master design](../15-deliverables/MASTER-TECHNICAL-DESIGN.md), module inventory and actual decision register {decisions}. Compare the exported commit/dirty-content baseline before relying on line numbers.
2. Use Windows desktop, Visual Studio Office/VSTO workload and .NET Framework 4.8 targeting pack for host work. Pure chart packages require Node (engine package declares >=24); locked AJV/TypeScript dependencies must already exist or be installed by the normal reviewed developer setup.
3. Understand three independent success levels: deterministic contracts/geometry tests, signed packaging integrity and actual interactive Office acceptance. A green portable suite does not prove VSTO registration, ribbon visibility or native save/reopen.
4. Locate UI routing in `ThisAddIn`, generated ribbon definitions/catalogs and context controller; find canonical codec/CustomXML attachment ownership before changing a feature. Select the existing compatible metadata namespace and version deliberately.
5. Run the appropriate existing test scripts in a developer session with presentations saved/closed for Office tests. This documentation task intentionally does not run scripts that would modify source/artifacts/configuration outside export.
6. Keep public display branding separate from internal COM IDs, namespaces, file paths and metadata identifiers. Preserve local library licenses and publisher trust rather than renaming identities without migration.

Build instructions and configuration are in [development setup](../11-build-deployment/development-setup.md). New production features need model validation, non-mutation rejection tests, native editing and save/reopen evidence. AI/service requirements must not be represented by fake returned content or endpoint-only readiness. {services}.''')
put('13-maintenance/debugging-guide.md','Debugging guide',f'''
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

Fixed-operation diagnostics intentionally avoid workbook contents. Catch boundaries should not be used as proof a whole mutation was atomic. `StartNewUndoEntry` groups native operations; it is not a transaction covering all COM calls. Use narrow source-specific integration tests. {event}; {run}; {lease}; {surface}.

COM failures may arise from the execution token (AppContainer), host focus, bitness, type-library registration, native workbook lifetime or busy Office calls. HRESULT alone is insufficient diagnosis. Never respond by disabling security tools, changing COM ACLs or force-terminating user Office processes. Historical activation evidence is environment-specific; rerun diagnostics only in an authorized developer session.''')
put('13-maintenance/extension-guide.md','Extension guide',f'''
## Add a chart or visual family

Choose among Excel-backed native chart creation, grouped editable shapes, procedural library layouts and faithful original slide insertion. These routes have different persistence/editability semantics. Define model constraints and capability rejection first. Add pure calculations/scene tests before Office mutations, then map native semantic bindings and preserve omission diagnostics. {renderer}; {surface}.

The native feasibility adapter verifies a canonical v2 clustered-column model against a rectangle scene. Extending TypeScript `buildScene` to another family alone does not extend that C# adapter. Do not silently reinterpret V3 files that are not compiled into current targets. {domain}; {ref('packages/chart-engine/src/scene-builder.ts','export async function')}.

## Add UI commands

Add a catalog/definition entry, callback route, enablement policy and actual implementation. Use the existing UI STA, guarded `Run`, selection identity and undo boundary. Include tests asserting the callback exists, the target is appropriate and unsupported selections fail before mutation. Keep manually chosen native formatting when canonical refresh policy requires it. {run}; {ref('apps/powerpoint-vsto/DemoRibbon.cs')}.

## Extend local library/data integration

PIIDX changes require explicit schema version and migration compatibility, parameterized statements, source fingerprint refresh, tested duplicate/search behavior and preview integrity. A new input source needs bounded parsing, explicit mapping and a source-change failure policy. Never treat a local profile as an authenticated connector. {schema}; {catalog}; {ref('apps/powerpoint-vsto/DataSources.cs')}.

## Add governed AI/services

Implement approved identity/network/data-policy decisions, transport client, request/response schemas, timeout/retry policy and action approval before wiring a ribbon command. Existing JSON schema is a proposed contract, not a running orchestration engine. Do not permit model text to invoke COM/code/SQL. {services}.''')
put('13-maintenance/operational-runbook.md','Operational runbook',f'''
**Audience:** interactive Windows support/developer staff. No server deployment or administrative action is performed by this export.

Before rollout, verify exact code/package hashes, architecture and publisher trust in the intended Windows user. Save/close presentations, use the normal installer prompt and then confirm NexWave ribbon loading. Retain extraction folders referenced by VSTO. Use the current package validation script rather than judging readiness from a copied filename. {ref('easySetup/Test-EasySetup.ps1')}.

For a managed chart failure, capture only fixed stage/error metadata and the owned test evidence directory. Preserve native objects and original local source files. If refresh sees a moved file or changed headers, require confirmed relinking; do not automatically guess replacements. For library failures, check source registration, source root permissions and exact schema/integrity validation before replacing a catalog. User source folders are authoritative; an index is derived local state. {ref('apps/powerpoint-vsto/MappedLinkOperations.cs')}; {catalog}.

Artifact cleanup is a separate reviewed operation. Preserve active deployment files, stable runner caches/receipts, signing material and scanner history. Historical source/evidence archives are not executable recovery sources. The documentation export is never a source-distribution/install bundle. {ref('scripts/Maintain-Artifacts.py')}.

Uninstall and rollback use VSTO/easySetup's implemented mechanisms. There is no enterprise background scheduler, central telemetry, RBAC service or documented HA/DR runtime to operate. Backup and retention of presentations, linked files and source libraries need local organizational policy; default JSON and PIIDX paths are not a substitute. Open operational decisions remain Q-005/Q-006/Q-007/Q-008/Q-012.''')
print('parent architecture and maintenance docs written')
