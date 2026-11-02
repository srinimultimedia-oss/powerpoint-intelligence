POWERPOINT INTELLIGENCE  |  ARCHITECTURE & FUNCTIONAL DESIGN  |  v1.0\
\
AI-Powered PowerPoint Intelligence Add-in\
\
Complete Architecture, Functional Design, API Contracts, Data Models, Security Controls, Acceptance Tests & Codex Implementation Playbook\
\
Architecture baseline: Windows PowerPoint • VSTO/.NET Framework 4.8 • On-premises services • Azure OpenAI • Native editable PowerPoint objects\
\
Document status: DEVELOPMENT-READY BASELINE WITH EXPLICIT OPEN DECISIONS. This is a proposed independent product design, not documentation of proprietary think-cell internals.\
\
Attribute\
\
Value\
\
Version\
\
1.0 / 21 October 2026\
\
Audience\
\
Solution architects, developers, QA engineers, security, DevOps, Codex agents\
\
Decision authority\
\
Product owner / enterprise architecture review\
\
Scope\
\
68 previously numbered core requirement entries + 30 AI requirements; 98 tracked IDs (includes 2 enabling requirements)\
\
Implementation strategy\
\
Vertical slices, deterministic core, explicit feature flags and evidence gates\
\
How to use this document\
\
Read Sections 1–7 before writing code. Section 8 specifies functional requirements. Sections 9–13 provide API and data contracts, threat controls, acceptance criteria, and Codex-ready delivery instructions. Each proposed feature is identified by an ID. Do not silently fill undecided parameters.\
\
1\. Evidence, decisions, assumptions and non-goals\
\
Tag\
\
Interpretation\
\
VERIFIED\
\
Supported by Microsoft documentation cited in Section 17\
\
DECIDED\
\
Explicitly chosen by the product owner\
\
PROPOSED\
\
Architecture recommendation; subject to spike and review\
\
OPEN\
\
Must be decided or experimentally verified before production implementation\
\
DECIDED: Windows PowerPoint is the first and primary host; Mac and Web are out of scope for v1.\
\
DECIDED: Business application, storage, integration and orchestration services run on-premises.\
\
DECIDED: Azure OpenAI is the sole AI model provider; its inference endpoint is Azure-hosted, not physically on-premises.\
\
DECIDED: Output prioritizes native editable PowerPoint shapes, charts, text and tables; a custom semantic document model retains full editability through the add-in.\
\
VERIFIED: VSTO add-ins run with .NET Framework; Microsoft does not support in-process modern .NET VSTO add-ins.\
\
VERIFIED: PowerPoint COM interop supports desktop object-model automation; Open XML SDK can manipulate PresentationML packages outside Office.\
\
PROPOSED: C# .NET Framework 4.8 VSTO add-in + separate local .NET 8+ worker + on-prem ASP.NET Core API + TypeScript calculation packages.\
\
OPEN: Exact Microsoft 365/Office build matrix, 32/64-bit distribution, Azure region/model deployment, licensing rules, identity provider, offline policy and data classification.\
\
2\. Business objectives and measurable outcomes\
\
Objective\
\
Outcome / KPI\
\
Editable analytical charts\
\
All supported charts reopen with original data and editable native components\
\
Data trust\
\
Zero untracked numeric changes; deterministic calculations and provenance\
\
AI productivity\
\
Natural-language actions produce previewable, validated structured commands\
\
Enterprise control\
\
No sensitive data leaves on-premises boundary except explicitly permitted Azure inference requests\
\
Reliability\
\
No known document corruption in supported test corpus; rollback on partial mutations\
\
Codex readiness\
\
Each development task has explicit files, contracts, tests and completion evidence\
\
3\. System context and deployment boundaries\
\
Windows desktop process: PowerPoint.exe loads the VSTO add-in (C#/.NET Framework 4.8). It owns Ribbon commands, selection handling, custom task pane hosting, and serialized COM mutations. Never invoke PowerPoint COM objects from arbitrary background threads.\
\
On-premises zone: API gateway, application backend, AI orchestration service, database, report queue, template/asset store, logs and audit trail. Optional per-user local worker handles computationally intensive chart calculations and task-pane UI bridging. Azure zone: approved Azure OpenAI deployment accessed over private connectivity and governed egress.\
\
Critical distinction: private Azure connectivity does not mean inference occurs on-premises. Security sign-off must cover Azure processing, residency, retention, and service configuration.\
\
Concept architecture reference\
\
\
\
Illustrative architecture image from the design discussion. The detailed decisions and boundaries in this document take precedence where the illustration is simplified.\
\
4\. Logical components and ownership\
\
Component\
\
Responsibility\
\
Location/runtime\
\
Hard boundary\
\
PowerPointHost\
\
VSTO, Ribbon, selection and events\
\
In-process .NET Framework 4.8\
\
PowerPoint COM; no network secrets\
\
TaskPane\
\
React editor and AI Copilot\
\
Embedded WebView2 host (feasibility spike)\
\
UI-only, no direct COM calls\
\
CommandCoordinator\
\
Commands, revisions, approval, rollback\
\
In-process C#\
\
Single-thread COM dispatcher\
\
ChartCore\
\
Semantic calculations and chart geometry\
\
Shared TypeScript package or compiled worker\
\
Pure deterministic functions\
\
LayoutCore\
\
Text metrics, constraints, labels, collisions\
\
Local worker\
\
Versioned scene graph\
\
NativeRenderer\
\
Native shapes, charts, tags, metadata\
\
PowerPoint COM adapter\
\
Reconcile stable semantic IDs\
\
DocumentStore\
\
Canonical model and migration\
\
Custom XML + shape tags (verify)\
\
Round-trip invariants\
\
DataConnectors\
\
Excel, CSV, API, database\
\
Local/on-prem connectors\
\
Read-only by default\
\
AIOrchestrator\
\
Intent, planning, tool calling\
\
On-prem ASP.NET Core\
\
Azure OpenAI only\
\
ModelGateway\
\
Redaction, policy, Azure transport\
\
On-prem ASP.NET Core\
\
Private networking; model allowlist\
\
Automation\
\
Templates, schedules, batch jobs\
\
On-prem workers\
\
No unattended desktop Office automation\
\
Governance\
\
RBAC, SSO, audit, telemetry\
\
On-prem services\
\
Policy enforcement at every boundary\
\
5\. Runtime interaction flows\
\
Create/edit chart\
\
Ribbon/Task pane → command DTO → schema/permission check → chart calculation → layout → preview → approved COM mutation → metadata save → audit event.\
\
AI edit\
\
Selection snapshot → redact/authorize → on-prem orchestrator → Azure OpenAI → structured action plan → deterministic validation → diff preview → approval → command bus → render → audit.\
\
Excel refresh\
\
Resolve workbook/range → fingerprint source → read values → normalize → reconcile → preview changes → update semantic model → render → persist link metadata.\
\
Batch report\
\
API request → authenticate/authorize → snapshot dataset/template versions → queue → server-side Open XML generation → validate PPTX → artifact store → signed/authorized download.\
\
Offline\
\
Use cached embedded data and local chart engine; queue optional sync metadata; AI calls unavailable unless corporate network connectivity to Azure path exists.\
\
6\. Native-editable output strategy\
\
Use native PowerPoint charts for simple chart families only when their editability and data semantics meet the contract; otherwise use editable PowerPoint shapes for bars, lines, labels, arrows and connectors.\
\
Canonical chart JSON/XML remains the source of truth; shapes are a projection. Native PowerPoint charts may have a separate embedded workbook requiring reconciliation.\
\
Assign a semantic object GUID and child role identifiers; maintain a mapping from semantic nodes to PowerPoint shape IDs/tags and validate after copy/paste, grouping and slide duplication.\
\
Store versioned metadata in supported custom XML parts and verify API behavior in a spike; shape tags are supplemental and should not be treated as guaranteed immutable identity.\
\
Never mutate the ZIP/PresentationML of an open presentation behind PowerPoint. Server-side Open XML operates on closed file copies only.\
\
For each chart, define a capability profile: native chart, native shape group, editability, unsupported features, and fallback.\
\
7\. Engineering contracts and repository layout\
\
/src\
&#x20; /PowerPoint.AddIn                 # VSTO .NET Framework 4.8\
&#x20; /PowerPoint.InteropAdapter       # COM access, STA command executor\
&#x20; /PowerPoint.TaskPane             # React/TypeScript, hosted WebView2\
&#x20; /Chart.Domain                    # versioned semantic models\
&#x20; /Chart.Calculation               # pure algorithms\
&#x20; /Chart.Layout                    # geometry, labels, constraints\
&#x20; /Chart.Renderers                 # scene graph to COM operations\
&#x20; /Data.Connectors                 # Excel, CSV, REST, database\
&#x20; /AI.Orchestrator                 # on-prem planning and tool gateway\
&#x20; /Enterprise.Api                  # ASP.NET Core REST\
&#x20; /Enterprise.Workers              # queue, reports, scheduling\
&#x20; /Enterprise.Persistence          # EF Core / PostgreSQL or SQL Server\
&#x20; /Shared.Contracts                # JSON Schema, OpenAPI, DTOs\
/tests/unit /tests/property /tests/contract /tests/ppt-integration\
/tests/visual /tests/security /tests/e2e\
/docs/adr /docs/feature-specs /docs/runbooks /docs/evidence\
\
interface ChartDocument {\
&#x20; schemaVersion: number; chartId: string; revision: number;\
&#x20; chartType: ChartType; dataset: DatasetRef;\
&#x20; axes: Axis[]; series: Series[]; annotations: Annotation[];\
&#x20; layout: Layout; style: Style; overrides: ManualOverride[];\
&#x20; source?: DataSourceRef; provenance: ProvenanceRecord[];\
}\
interface CommandEnvelope {\
&#x20; commandId: string; idempotencyKey: string;\
&#x20; targetChartId?: string; expectedRevision?: number;\
&#x20; actionType: string; arguments: unknown;\
&#x20; actorId: string; approvalId?: string;\
}\
interface RenderResult {\
&#x20; semanticScene: SceneNode[]; diagnostics: Diagnostic[];\
&#x20; renderHash: string; modelRevision: number;\
}\
\
Design rules: define JSON Schema and C# DTO equivalents; no untyped model-generated commands at execution boundary; use optimistic revision checks; serialize COM writes on the PowerPoint main STA; all layout units use points with explicit conversions.\
\
8\. Functional requirements and primary acceptance criteria\
\
The following entries are actionable requirements. Each acceptance criterion is a minimum test, not an exhaustive test suite. All requirements additionally inherit the global negative, security, persistence and accessibility tests in Section 12.\
\
8.1 Chart families\
\
ID / Feature\
\
Required behavior\
\
Primary acceptance criterion\
\
CHT-001 Clustered column/bar\
\
Render grouped series with vertical/horizontal category axes.\
\
3 series × 4 categories yields 12 bars with correct ordering.\
\
CHT-002 100% stacked\
\
Normalize positive category totals to 100%; define signed-data policy.\
\
20/30/50 renders 20%/30%/50%; zero total is handled.\
\
CHT-003 Combination charts\
\
Per-series type and axis assignment.\
\
Bars + line preserve values and distinct axis scales.\
\
CHT-004 Mekko / Marimekko\
\
Weighted variable widths and stacked proportions.\
\
10/30/60 widths map to 10/30/60% of plot width.\
\
CHT-005 Scatter and bubble\
\
Numeric XY coordinates and area-scaled bubble sizes.\
\
Coordinates are correct; bubble area follows size field.\
\
CHT-006 Histogram and Pareto\
\
Numeric binning; descending Pareto and cumulative percentage.\
\
Bin counts match fixture; Pareto ends at 100%.\
\
CHT-007 Small multiples\
\
Facet repeated panels with shared/independent axis options.\
\
Six facets share identical scale in shared-axis mode.\
\
CHT-008 Radar and heatmap\
\
Radial series and matrix color mapping.\
\
Radial values and heatmap thresholds match fixtures.\
\
CHT-009 Gantt and timeline\
\
Date ranges, milestones, dependencies, progress.\
\
10 tasks and dependencies remain anchored after date edit.\
\
CHT-010 Pie and doughnut\
\
Nonnegative proportional sectors and editable labels.\
\
25/25/50 maps to 90/90/180 degrees.\
\
CHT-011 Waterfall and bridge\
\
Cumulative deltas, explicit totals/subtotals, connectors.\
\
100 + 40 − 25 = 115; total not double-counted.\
\
CHT-012 Line and area\
\
Ordered points, gap policy, baseline, interpolation.\
\
Missing point follows selected gap policy.\
\
CHT-013 Stacked column/bar\
\
Positive and negative stacks independently around zero.\
\
Mixed-sign series do not overlap across zero.\
\
8.2 Chart intelligence\
\
ID / Feature\
\
Required behavior\
\
Primary acceptance criterion\
\
INT-001 Automatic data labels\
\
Format and place semantic data labels.\
\
Values refresh and empty-point policy holds.\
\
INT-002 Dynamic totals/subtotals\
\
Recompute aggregates from current data.\
\
Change 10→15 increases dependent total by 5.\
\
INT-003 Absolute/percentage difference arrows\
\
Compute delta and percent against selected baseline.\
\
80→100 shows +20 and +25%; zero denominator flagged.\
\
INT-004 Legends and series labels\
\
Maintain series identity, order and direct labels.\
\
Rename series updates all bound labels.\
\
INT-005 Error indicators\
\
Represent absolute, relative or confidence error ranges.\
\
50 ± 5 yields 45–55.\
\
INT-006 Annotations and connectors\
\
Anchor callouts to semantic points.\
\
Resize retains attachment.\
\
INT-007 Collision avoidance\
\
Solve overlaps with priority/fallback policy.\
\
Dense fixture has no prohibited overlaps.\
\
INT-008 CAGR arrows\
\
Compute CAGR with defined time intervals and domain.\
\
100→121 over 2 years yields 10%.\
\
INT-009 Value axes and breaks\
\
Linear/log/time scales and marked discontinuities.\
\
Log rejects nonpositive values; break visibly marked.\
\
INT-010 Custom number formats\
\
Locale-aware formatting without changing values.\
\
1234.5 and negative counterpart display correctly.\
\
INT-011 Data-driven colors\
\
Color by sign, threshold or category.\
\
Threshold boundary follows inclusive rule.\
\
8.3 Editing and layout\
\
ID / Feature\
\
Required behavior\
\
Primary acceptance criterion\
\
EDT-001 Drag handles\
\
Move/resize editable objects within bounds.\
\
Corner drag honors min dimensions.\
\
EDT-002 Snap and align\
\
Snap to guides and object anchors.\
\
Within tolerance snaps exactly.\
\
EDT-003 Multi-selection\
\
Batch-edit compatible objects.\
\
Theme applied to three charts preserves data.\
\
EDT-004 Keyboard shortcuts\
\
Configurable, focus-aware shortcuts.\
\
Editor typing not intercepted.\
\
EDT-005 Group and ungroup\
\
Maintain semantic identity where supported.\
\
Group/reopen/ungroup preserves mapping.\
\
EDT-006 Chart conversion\
\
Convert compatible semantics with warnings.\
\
Clustered→stacked preserves raw values.\
\
EDT-007 Resize/reflow\
\
Recalculate geometry and labels on resize.\
\
40% width reduction does not clip mandatory content.\
\
EDT-008 Distribute objects\
\
Equalize horizontal/vertical spacing.\
\
Four shapes have equal gaps.\
\
EDT-009 Context menus\
\
Show context/permission-aware actions.\
\
Label right-click offers label actions.\
\
EDT-010 Undo/redo\
\
Transactional command history and native undo policy.\
\
Edit/resize undo twice and redo twice restores states.\
\
EDT-011 Theme-aware formatting\
\
Inherit theme with manual overrides.\
\
Theme changes inherited colors only.\
\
EDT-012 Selection synchronization\
\
Reflect current selection in task pane.\
\
Selecting another chart shows correct properties.\
\
EDT-013 Manual override preservation\
\
Preserve compatible custom edits after reflow.\
\
Refresh retains compatible label overrides.\
\
8.4 Data integration\
\
ID / Feature\
\
Required behavior\
\
Primary acceptance criterion\
\
DAT-001 Embedded data editor\
\
Spreadsheet-like typed cell editing.\
\
Edit/save/reopen retains values.\
\
DAT-002 Link refresh\
\
Detect changes and update linked charts.\
\
Excel source change refreshes chart and timestamp.\
\
DAT-003 Formula handling\
\
Parse allowed formulas, dependencies, errors.\
\
SUM recalculates; circular reference flagged.\
\
DAT-004 Database/API datasets\
\
Authenticated approved read-only connectors.\
\
Bad schema/token fails safely.\
\
DAT-005 Data provenance\
\
Track source, transforms, timestamps and hashes.\
\
Chart inspector shows lineage.\
\
DAT-006 Validation and reconciliation\
\
Compare typed data and aggregate totals.\
\
100 source vs 95 transformed reports difference 5.\
\
DAT-007 Scheduled report data\
\
Retrieve snapshots for scheduled reports.\
\
Run records dataset version and outcome.\
\
DAT-008 CSV import\
\
Parse quotes, delimiters, Unicode, locales.\
\
Quoted commas and UTF-8 preserved.\
\
DAT-009 Broken link recovery\
\
Relink relocated sources without chart recreation.\
\
Moved workbook relinks and refreshes.\
\
DAT-010 Excel range linking\
\
Named ranges, sheets and structured tables.\
\
Table growth follows declared policy.\
\
8.5 Slide productivity\
\
ID / Feature\
\
Required behavior\
\
Primary acceptance criterion\
\
SLD-001 Agenda and section slides\
\
Generate sections from outline.\
\
Five sections yield five references.\
\
SLD-002 Harvey balls\
\
Editable fraction-filled circles.\
\
75% renders and persists.\
\
SLD-003 Process diagrams\
\
Editable nodes and connected edges.\
\
Five-step process retains connectors.\
\
SLD-004 Slide templates\
\
Versioned placeholders and bindings.\
\
Template preserves placeholder IDs.\
\
SLD-005 Batch formatting\
\
Apply styles to scoped slides.\
\
20 titles updated, unrelated text unchanged.\
\
SLD-006 Presentation QA\
\
Find overflow, broken links, missing data.\
\
Fixture flags all three issues.\
\
SLD-007 Slide library\
\
Search, version and insert slide assets.\
\
Inserted slide records asset version.\
\
SLD-008 Corporate themes\
\
Brand palettes, fonts and locks.\
\
Locked objects remain unchanged.\
\
SLD-009 Icons and pictograms\
\
Insert accessible scalable icons.\
\
Resizing preserves aspect and alt description.\
\
SLD-010 Checkboxes\
\
Editable tri-state visual controls.\
\
Toggle and reopen retains state.\
\
SLD-011 Tables\
\
Editable data-bound tables and totals.\
\
5×4 table recalculates totals.\
\
8.6 Enterprise features\
\
ID / Feature\
\
Required behavior\
\
Primary acceptance criterion\
\
ENT-001 Template-based generation\
\
Bind data into versioned templates.\
\
Five placeholders all populated.\
\
ENT-002 REST API\
\
Authenticated report/template/data endpoints.\
\
Bad payload returns structured 4xx.\
\
ENT-003 Role-based access\
\
Authorize actions and artifacts.\
\
Viewer cannot publish or bypass API.\
\
ENT-004 Deployment policies\
\
Enforce managed configurations.\
\
Unapproved connector blocked.\
\
ENT-005 Offline mode\
\
Local editing with online feature degradation.\
\
Network loss does not block embedded edit.\
\
ENT-006 Licensing and updates\
\
Signed updates and entitlement checks.\
\
Tampered update rejected.\
\
ENT-007 Telemetry and diagnostics\
\
Privacy-filtered performance and errors.\
\
No dataset contents in error telemetry.\
\
ENT-008 SSO integration\
\
Enterprise identity and token validation.\
\
Expired token requires reauthentication.\
\
ENT-009 Audit history\
\
Immutable/tamper-evident significant actions.\
\
Publish action records actor and outcome.\
\
ENT-010 Batch presentation generation\
\
Isolated queued PPTX jobs.\
\
100 jobs individually traceable.\
\
8.7 AI requirements: planning, analysis and governance\
\
ID / Feature\
\
Behavior contract\
\
Acceptance criterion\
\
AI-001 Text-to-chart\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Valid prompt generates editable chart with exact values\
\
AI-002 Chart recommendation\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Recommendation cites field types and chart suitability\
\
AI-003 Conversational chart editing\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Only intended formatting changes after approval\
\
AI-004 AI chart conversion\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Incompatible conversion leaves original unchanged\
\
AI-005 Smart label optimization\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Label suggestion passes deterministic collision solver\
\
AI-006 Chart explanation\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Explanations separate fact, calculation and inference\
\
AI-007 Natural-language data mapping\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Ambiguous field names request clarification\
\
AI-008 Data quality detection\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Nulls/outliers/invalid dates flagged by rules\
\
AI-009 Natural-language dataset querying\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Only approved read-only query tools invoked\
\
AI-010 Trend and variance analysis\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Reference totals and changes match independent oracle\
\
AI-011 Forecast assistance\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Forecast includes method, horizon and uncertainty\
\
AI-012 Data reconciliation assistant\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Mismatch traced to source/transform if possible\
\
AI-013 Text-to-slide\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Generated slide contains editable objects\
\
AI-014 Smart slide layout\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Locked content remains unchanged\
\
AI-015 Corporate style assistant\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Only approved brand styles used\
\
AI-016 Slide redesign\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Preview and rollback available\
\
AI-017 Narrative/storyline builder\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Unsupported claims flagged, not invented\
\
AI-018 Editable diagram generation\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Nodes and connectors remain editable\
\
AI-019 Executive summary\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Every number traces to data/slide\
\
AI-020 Insight generation\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
No unsupported statistical assertion\
\
AI-021 Presentation critique\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
QA findings reference object locations\
\
AI-022 Speaker notes\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
No invented slide metrics\
\
AI-023 Audience adaptation\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Rewrites preserve numerical facts\
\
AI-024 Source verification\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Unavailable evidence labeled unverified\
\
AI-025 AI report builder\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Approved template and exact data used\
\
AI-026 Multi-step task planning\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Failed dependent action stops/rolls back\
\
AI-027 Natural-language automation\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Schedule/destination confirmed before activation\
\
AI-028 Change-impact analysis\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Affected slide/chart/source revisions identified\
\
AI-029 Sensitive data protection\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Restricted data blocked or redacted before Azure call\
\
AI-030 AI action audit and approval\
\
Model proposes structured, permission-scoped operations; deterministic validation applies.\
\
Approval and executed commands fully audited\
\
9\. Data design and storage\
\
PROPOSED relational database: PostgreSQL or SQL Server, selected by enterprise standards. Every table includes tenant/business-unit scope where applicable, UTC timestamps, created/updated metadata and concurrency tokens. Do not store API keys or raw prompts with sensitive content in business tables.\
\
Entity\
\
Key fields\
\
Constraints\
\
Presentation\
\
id, owner_id, title, current_revision, classification\
\
revision monotonic\
\
Chart\
\
id, presentation_id, slide_guid, schema_version, revision, model_json\
\
unique semantic ID\
\
ChartObjectMap\
\
chart_id, semantic_node_id, shape_tag, last_seen_shape_id\
\
reconciliation on copy/duplicate\
\
Dataset\
\
id, schema_json, source_type, fingerprint, revision\
\
immutable snapshot revisions\
\
DataLink\
\
id, dataset_id, source_locator, refresh_policy, state\
\
source credentials referenced by secret ID\
\
Template\
\
id, version, content_hash, approval_state\
\
published versions immutable\
\
ReportJob\
\
id, template_version, dataset_snapshot, state, result_uri\
\
idempotency key unique\
\
AIActionPlan\
\
id, actor_id, model_deployment, prompt_hash, plan_json, state\
\
approval and revision checks\
\
Approval\
\
id, plan_id, approver_id, decision, expires_at\
\
single-use approval token\
\
AuditEvent\
\
id, actor, action, object_id, correlation_id, outcome, hash_chain\
\
append-only access\
\
Policy\
\
id, scope, version, settings_json\
\
signed or privileged updates\
\
Asset\
\
id, version, media_type, content_hash, storage_uri\
\
antivirus and file validation\
\
Canonical chart storage: serialized versioned semantic chart model in the PPTX, with a database copy only when enterprise synchronization is enabled. For disconnected presentations, the PPTX is self-contained; cloud/on-prem metadata is never required to open the file.\
\
10\. API contracts and status semantics\
\
Endpoint\
\
Method\
\
Access\
\
Contract\
\
/api/v1/ai/plans\
\
POST\
\
Editor\
\
Create proposed structured action plan; no mutation\
\
/api/v1/ai/plans/{id}/validate\
\
POST\
\
Editor\
\
Revalidate revisions, policy, schema\
\
/api/v1/ai/plans/{id}/approve\
\
POST\
\
Approver\
\
Approve bounded action hash\
\
/api/v1/datasets/validate\
\
POST\
\
Editor\
\
Typed schema, errors, warnings\
\
/api/v1/charts/render\
\
POST\
\
Editor\
\
Scene graph preview; never execute COM server-side\
\
/api/v1/templates\
\
GET/POST\
\
Viewer/Publisher\
\
List or publish versioned templates\
\
/api/v1/reports\
\
POST\
\
Generator\
\
Create idempotent report job\
\
/api/v1/reports/{id}\
\
GET\
\
Authorized owner\
\
Status, diagnostics, result link\
\
/api/v1/batches\
\
POST\
\
Generator\
\
Create isolated batch jobs\
\
/api/v1/audit-events\
\
GET\
\
Auditor\
\
Filtered, authorized audit records\
\
POST /api/v1/ai/plans\
{\
&#x20; "requestId":"uuid", "presentationId":"uuid",\
&#x20; "selection":{"slideIds":["slide-guid"],"chartIds":["chart-guid"]},\
&#x20; "instruction":"Make negative bars red",\
&#x20; "expectedRevisions":{"chart-guid":7},\
&#x20; "contextPolicy":"selection-only"\
}\
200 OK\
{\
&#x20; "planId":"uuid", "state":"proposed",\
&#x20; "actions":[{"type":"SetSeriesColor","targetId":"chart-guid", "args":{"rule":"negative","color":"#C44A4A"}}],\
&#x20; "requiresApproval":true, "warnings":[], "previewToken":"opaque"\
}\
\
API errors: 400 invalid request, 401 unauthenticated, 403 forbidden, 404 not found, 409 revision conflict, 413 payload too large, 422 invalid chart semantics, 429 quota, 502 model dependency failure, 503 unavailable. Responses include correlationId and machine-readable error code; never leak tokens or raw source data.\
\
11\. AI orchestration and Azure OpenAI security\
\
AI plans, deterministic executor acts. The model can recommend chart types, summarize approved content and emit typed action proposals; it must not execute arbitrary SQL, shell commands, VBA, or unrestricted COM calls. On-prem orchestrator supplies only authorized, minimized context. Validate all tool arguments against versioned schemas; prohibit instruction following from retrieved slide or spreadsheet content.\
\
Use enterprise-approved Azure OpenAI deployment and supported API version; configure via environment/secret manager, never hardcode endpoint/model/version.\
\
Network: on-prem gateway → enterprise firewall/proxy → private VPN/ExpressRoute → Azure VNet private endpoint → Azure OpenAI. Confirm DNS, routing, firewall, region and fail-closed behavior.\
\
Identity: prefer managed identity or enterprise service principal with certificate/secure secret storage as supported by deployment; no API keys in add-in binaries.\
\
Governance: enforce classification labels, consent/approval, prompt redaction, per-role tool allowlist, max tokens, cost budgets, retention controls and explicit model deployment version.\
\
AI tool output is untrusted until parsed and validated. Reject unknown action types, forbidden targets, mismatched revisions and nonfinite numbers.\
\
High-impact changes require user approval bound to exact plan hash, target IDs and expected revisions; approval expiry and one-time use are mandatory.\
\
Offline AI is unavailable unless a separately approved local inference capability is commissioned; local chart editing must remain functional.\
\
12\. Verification strategy, quality gates and negative tests\
\
Suite\
\
Required coverage / oracle\
\
Unit and property\
\
Calculation invariants, nonfinite inputs, mixed signs, zero totals, scaling, rounding\
\
Contract\
\
JSON Schema, OpenAPI, DTO compatibility, version migrations\
\
PPTX round trip\
\
Create→save→close→reopen→edit; copied charts, duplicates, grouped shapes\
\
COM integration\
\
STA behavior, selection changes, cancellation, process shutdown, Office build matrix\
\
Visual regression\
\
Approved fonts/DPI/theme and measured geometric tolerances\
\
AI safety\
\
Prompt injection in spreadsheet cell, invalid JSON, unauthorized actions, stale revisions\
\
Data integration\
\
Excel renamed/moved, CSV formula injection, invalid API schema, stale token\
\
Security\
\
SSO/RBAC bypass, secret leakage, tenant isolation, malicious template/ZIP\
\
Batch generation\
\
Concurrent jobs, retries, idempotency, isolation, corrupted template\
\
Accessibility\
\
Keyboard navigation, focus, contrast, chart alt descriptions\
\
Cross-feature E2E fixtures: E2E-001 create/edit/reopen; E2E-002 Excel refresh+undo; E2E-003 conversion round-trip; E2E-004 theme override; E2E-005 100-point labels; E2E-006 20-report template generation; E2E-007 moved workbook recovery; E2E-008 offline editing; E2E-009 copy across presentations; E2E-010 malformed metadata; E2E-011 native PowerPoint edit mixed with add-in; E2E-012 100-report batch with 3 invalid datasets.\
\
Release gates (PROPOSED): zero known data-corruption defects; 100% approved PPTX corpus save/reopen; all P0 acceptance tests green; numerical reference fixtures exact within declared tolerance; security threat model signed off; P95 preview <250 ms for defined small chart hardware; native PowerPoint commit P95 <1 second for small chart; failures rollback or recover to last known good state.\
\
13\. Codex implementation playbook — executable development order\
\
Instruction to coding agents: work in bounded pull requests; do not invent unavailable APIs or mark untested COM behavior as complete. Produce compilable code, unit tests, reproducible fixtures and a concise evidence report for every slice. Ask for product decisions when an OPEN item blocks the task.\
\
Slice\
\
Scope\
\
Codex task\
\
Exit evidence\
\
S0\
\
Environment and proof of concept\
\
Create .NET Framework 4.8 PowerPoint VSTO project; ribbon button; add native rectangle; save/reopen; verify build in Windows Office VM.\
\
VSTO add-in loads and edits one slide; screenshot and integration test log.\
\
S1\
\
Domain and contracts\
\
Create JSON Schema, DTOs, chart semantic IDs, versioning, validator and migrations.\
\
Golden serialization round-trip fixtures pass.\
\
S2\
\
Chart calculation engine\
\
Implement clustered, stacked, 100%, line, pie; pure algorithms and fixtures.\
\
All core numerical unit/property tests pass.\
\
S3\
\
Layout and renderer\
\
Create scene graph, text metrics, native shapes renderer, object reconciliation.\
\
PowerPoint reopen retains semantic mapping.\
\
S4\
\
Editor and commands\
\
React task pane, data grid, selection sync, undo transactions and diff preview.\
\
User edits and native host selection stay consistent.\
\
S5\
\
Advanced chart intelligence\
\
Waterfall, totals, labels, axes, CAGR/differences, collision solving.\
\
Numerical fixtures and visual snapshots pass.\
\
S6\
\
Data integration\
\
CSV/Excel links, fingerprints, reconciliation and link recovery.\
\
Moved workbook, invalid CSV and refresh cases pass.\
\
S7\
\
AI vertical slice\
\
On-prem AI gateway, Azure OpenAI, typed SetSeriesColor plan, approval, executor.\
\
No document mutation before approval; injection test blocked.\
\
S8\
\
Advanced families and slides\
\
Mekko, Gantt, scatter/bubble, heatmap, radar, small multiples, tables/templates.\
\
Each requirement has an approved fixture and E2E test.\
\
S9\
\
Enterprise automation\
\
Template registry, Open XML report generation, queue, batch, RBAC/SSO/audit.\
\
100 isolated report jobs with status and traceability.\
\
S10\
\
Production hardening\
\
Installer, code signing, deployment policy, telemetry, backup/restore, documentation.\
\
Security, operations and UAT sign-off.\
\
14\. Example Codex issue template\
\
TASK: S2-CHT-001 clustered column geometry\
PRECONDITIONS: domain schema v1 and fixture runner merged.\
FILES: /src/Chart.Calculation/\* /tests/unit/ClusteredColumnTests.\*\
INPUT: typed dataset with 3 series and 4 categories; width/height in pt.\
OUTPUT: semantic scene graph; no PowerPoint COM calls.\
RULES: finite values only; category order stable; positive/negative baseline; deterministic IDs.\
TESTS: 12 bars, known coordinates, empty categories, negative values, invalid data.\
DONE: compile, tests pass, snapshot approved, documentation and changelog updated.\
BLOCKERS: ask product owner if signed-value normalization policy is unspecified.\
\
15\. Milestones, staffing, dependencies and risk register\
\
Phase\
\
Indicative timing\
\
Deliverable\
\
M0\
\
Weeks 1–3\
\
COM/VSTO and metadata feasibility gate\
\
M1\
\
Weeks 4–8\
\
Foundation, schema, basic editable native chart\
\
M2\
\
Weeks 9–16\
\
Core charts, labels, task pane, embedded data\
\
M3\
\
Weeks 17–28\
\
Advanced charts, layout, Excel links\
\
M4\
\
Weeks 29–38\
\
AI chart copilot, enterprise security and governance\
\
M5\
\
Weeks 39–52+\
\
Templates, automation, batch, QA, production hardening\
\
Planning estimate only: a complete 98-requirement platform can exceed 52 weeks depending on team size, Office compatibility and advanced chart fidelity. Suggested team: 1 architect, 2 Windows/COM developers, 2 chart/layout developers, 1 frontend developer, 1 backend/AI developer, 1 QA automation engineer, fractional DevOps/security/UX.\
\
Risk\
\
Impact\
\
Mitigation / evidence gate\
\
COM thread affinity / Office crashes\
\
Critical\
\
Single-thread dispatcher; VM integration tests\
\
Undo behavior mismatch\
\
High\
\
Early feasibility spike; document explicit supported semantics\
\
Shape identity after copy/group\
\
High\
\
Semantic IDs + reconciliation + duplicate tests\
\
Font metrics and label fidelity\
\
High\
\
Reference font environment; collision snapshots\
\
Azure connectivity from on-prem\
\
High\
\
Network/private DNS/security spike before AI coding\
\
Model hallucinations\
\
Critical\
\
Typed plans, deterministic calculations, approvals\
\
Unattended Office automation on server\
\
Critical\
\
Use Open XML on server, not PowerPoint COM\
\
PPTX corruption during metadata updates\
\
Critical\
\
Never ZIP-mutate open file; closed-copy package validation\
\
Excel link variability\
\
High\
\
Fingerprint/relink workflows; matrix of workbook states\
\
Licensing and IP boundaries\
\
High\
\
Independent implementation and legal review\
\
16\. Open decisions / questions requiring product owner or infrastructure sign-off\
\
ID\
\
Decision\
\
Why needed\
\
OD-01\
\
Supported Microsoft 365 PowerPoint builds and 32/64-bit Office?\
\
Blocks installer and COM compatibility matrix\
\
OD-02\
\
Is Microsoft Entra ID synchronized with on-prem AD, or another SSO provider?\
\
Blocks SSO implementation\
\
OD-03\
\
Which on-prem database is approved: SQL Server or PostgreSQL?\
\
Blocks persistence choice\
\
OD-04\
\
Azure OpenAI region, deployment/model, quotas, approved private connectivity?\
\
Blocks AI deployment\
\
OD-05\
\
Which corporate data classes may be sent to Azure OpenAI?\
\
Blocks production AI access\
\
OD-06\
\
Must every chart be editable using native PowerPoint chart controls, or is editable native shape composition acceptable?\
\
Blocks renderer acceptance\
\
OD-07\
\
Expected scale: users, concurrent report jobs, maximum slides/chart points?\
\
Blocks sizing and SLOs\
\
OD-08\
\
Offline license grace period and offline AI requirement?\
\
Blocks licensing/offline design\
\
OD-09\
\
Supported Excel source types: local files, network shares, SharePoint, OneDrive?\
\
Blocks link connector scope\
\
OD-10\
\
Brand standards, fonts, templates, icon licenses and slide library storage?\
\
Blocks corporate design assets\
\
OD-11\
\
Retention for prompts, telemetry, reports and audit records?\
\
Blocks privacy and storage policies\
\
OD-12\
\
Which numeric rules are approved for signed 100% stacks, negative CAGR and empty Mekko categories?\
\
Blocks chart edge-case semantics\
\
Codex operating rule: when a decision is OPEN, implement a configurable interface or stop at a testable stub; never silently choose a production policy.\
\
17\. Authoritative references and evidence ledger\
\
Microsoft: VSTO Add-ins and .NET Framework — https\://learn.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio?view=vs-2022\
\
Microsoft: VSTO Runtime Lifecycle — https\://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime?view=visualstudio\
\
Microsoft: PowerPoint VSTO walkthrough — https\://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-powerpoint?view=visualstudio\
\
Microsoft: Architecture of VSTO Add-ins — https\://learn.microsoft.com/en-us/visualstudio/vsto/architecture-of-vsto-add-ins?view=visualstudio\
\
Microsoft: Open XML SDK overview — https\://learn.microsoft.com/en-us/office/open-xml/about-the-open-xml-sdk\
\
Microsoft: PresentationML slides — https\://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-presentation-slides\
\
Microsoft: Custom XML parts — https\://learn.microsoft.com/en-us/visualstudio/vsto/custom-xml-parts-overview\
\
Microsoft: Azure AI private endpoints — https\://learn.microsoft.com/en-us/azure/ai-services/cognitive-services-virtual-networks\
\
Microsoft: Azure OpenAI network security (classic reference) — https\://learn.microsoft.com/en-us/azure/foundry-classic/openai/how-to/network?view=foundry-classic\
\
Evidence note: documentation verifies platform-level capabilities, not the performance or correctness of this proposed implementation. All functional algorithms, data schemas, acceptance tests and timelines in this document are engineering proposals until validated by implementation and QA.\
\
18\. Immediate development authorization checklist\
\
Approve OD-01, OD-04, OD-05 and OD-06 before a production architecture freeze.\
\
Create Windows development VM with approved Office build, Visual Studio Office workload and .NET Framework 4.8.\
\
Initialize monorepo and branch protection; publish JSON schemas and ADR-001 (VSTO boundary).\
\
Build S0 proof of concept and capture objective save/reopen evidence.\
\
Begin S1/S2 deterministic chart engine in parallel with COM feasibility work.\
\
Do not start bulk AI agents or batch report development until command bus, audit, permissions and persistence gates are green.