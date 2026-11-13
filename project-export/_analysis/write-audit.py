from pathlib import Path
E=Path('project-export')
def put(path,title,s):(E/path).write_text('# '+title+'\n\n'+s.strip()+'\n',encoding='utf-8')
put('14-audit/technical-debt-register.md','Technical debt register', '''These are source-review findings, not a remediation commit or a fresh Checkmarx result. Severity is qualitative maintenance/security impact; no exploit is asserted unless demonstrated.

| ID / priority | Finding and evidence | Consequence | Recommended engineering response |
|---|---|---|---|
| TD-01 High | V3 chart classes exist but both active projects omit them (`src/Chart.Domain/Chart.Domain.csproj:14`; `apps/powerpoint-vsto/PowerPointIntelligence.csproj:45`) | Engineers can mistake authored options for persisted production capability | Deliberate compile/migration decision, direct V3 contract/compatibility tests; do not silently add to build |
| TD-02 High | Active source/imported-original/procedural520 modes coexist (`apps/powerpoint-vsto/TemplateLibrary.cs:1`; `src/Library.Core/PresentationSource.cs:1`) | Library counts/design provenance/editing semantics can be conflated | Keep provenance/capability UI and tests explicit; preserve license boundaries |
| TD-03 High | TestS1 fallback NuGet download before independent identity verification (`scripts/Test-S1.ps1:18`) | Supply-chain coverage differs across test/build paths | Reuse a pinned approved digest/signature path, fail closed and tamper-test |
| TD-04 High | Scanner collector targets scene/template runners, not PIIDX (`scripts/Collect-SceneThreatEvidence.ps1:23`) | Historical absence of a PIIDX match was insufficient evidence | Scope/positive-control actual PIIDX file/hash without restoring detected binaries |
| TD-05 Medium | Locale-dependent fingerprint ordering (`scripts/Source-Fingerprint.ps1:1`; `easySetup/Common.ps1:14`) | NLS/ICU can mark unchanged package stale | Ordinal shared canonical recipe plus PS5/PS7 vectors and new signed release |
| TD-06 Medium | Broad UI catch and dense single-line implementations (`apps/powerpoint-vsto/ThisAddIn.cs:133`) | Root cause visibility and safe review become difficult | Keep user-safe status and improve fixed structured diagnostics; focused refactors with native regression |
| TD-07 Medium | Original native adapter is narrower than scene DTO/TS engine (`src/Chart.Renderers/NativeRenderingCore.cs:18`) | Type presence can imply unsupported native geometry | Explicit capability matrix and preflight rejection; add families one tested path at a time |
| TD-08 Medium | MJS utility calculations lack consistent result/accumulation finite guards (`packages/chart-core/src/calculations.mjs:1`) | Extreme finite inputs can overflow offline results | Negative overflow tests and explicit error/result policy; no production native impact assumed |
| TD-09 Medium | Source reader bounds entries/XML but no aggregate decompression budget (`src/Library.Core/PresentationSource.cs:1`) | Large packages can still consume work across entries | Aggregate byte/count/time/cancellation limits and stress fixtures |
| TD-10 Medium | Related installer/package readers retain implicit XML casts (`scripts/Install-Local.ps1:21`; `scripts/Test-DemoPackage.ps1:27`) | Explicit safe XML policy is inconsistent | Consolidate bounded DTD-prohibited readers; prove no app/config regression |
| TD-11 Medium | Native tests are environment/foreground/token dependent (`tests/s0/ContextPropertiesOfficeTests.cs:1`) | Infrastructure prerequisites can be mislabeled feature failure/pass | Preserve BLOCKED distinction, owned cleanup and per-feature evidence |
| TD-12 Medium | Enterprise requirements and current starter README differ from implemented desktop breadth (`README.md:1`; `services/ai-gateway/README.md:1`) | Onboarding/acceptance confusion | Maintain source-backed traceability and separate roadmap from shipping subsets |

Duplicated root chart-core.yml is not an active workflow; prefer one authoritative CI definition. Symbolic declaration extraction has known limitations with compressed statements; use source bodies for consequential edits. No claim of exhaustive dead-code elimination or automated call-graph reachability is made.''')
put('14-audit/implementation-gap-analysis.md','Implementation gap analysis', '''## Implemented source boundaries

VSTO lifecycle/ribbon/taskpane/context properties; bounded S0/v2 metadata; linked local data and explicit Excel mapping; native chart controls/tables/elements/agendas; procedural and original-template library paths; PIIDX metadata/search/cache; deterministic offlineJS/TS rendering; dual-bitness signed setup recipe. These are source facts, not every-path runtime passes. Feature catalog190 entries are172routes+18cross-cutting capabilities;103 descriptive requirement records remain separate.

## Partial or narrower than requirements

- Native semantic renderer: clustered-column rectangle feasibility only, with axes/legend omission diagnostics. General scene DTO geometries and six TS families do not imply all native chart families.
- Mapped refresh: uniform primary-axis category-chart subset; unsupported XY/bubble/pie/doughnut/secondary-axis/mixed cases are rejected rather than transparently synchronized.
- Template fidelity: source import preserves original formatting path; historical16-original sample is not exhaustive497-slide native acceptance. Procedural520 layouts are a distinct generated mode.
- Context properties: selection/foreground/applicability-limited controls, not every possible chart/object/DPI combination.
- Compatibility: requested product profiles and architecture checks, not a certified Office estate matrix.
- V3: source definitions exist, not integrated into active builds or migration acceptance.

## Planned or not implemented

On-prem gateway execution, live Azure OpenAI client/prompt orchestration/retry/token budgets, enterprise REST service, RBAC/SSO/tenant authorization, centralized audit/telemetry, enterprise report scheduling, production database/HA/DR and license/grace enforcement have no running implementation in inspected source. A local profile or requirements schema does not satisfy them. `services/ai-gateway/README.md:1`; `apps/powerpoint-vsto/ServiceConnections.cs:1`; `docs/DECISIONS.md:1`.

## Traceability and evidence gaps

Command-to-source documentation is broader than executed integration coverage. Most native demo acceptance tests compile production static implementations into an owned test runner; only dedicated installed-marker/context observation paths test parts of the installed add-in. Do not equate standalone COM editing success with complete installed-ribbon acceptance. Test existence/mocked planners and source-marker assertions are not feature acceptance. Details: [coverage gaps](../12-testing/coverage-gap-analysis.md), [feature status](../05-features/feature-status-matrix.md), [security gap register](../09-security/security-gap-register.md).

No missing feature was implemented by this assignment. Recommendations are maintenance backlog items requiring their own decisions and tests.''')
put('14-audit/architecture-risk-register.md','Architecture risk register', '''| Risk / assessed priority | Actual boundary | Impact and mitigation |
|---|---|---|
| AR-01 High — Office UI/COM coupling | Most feature logic mutates native objects on interactive UI STA | Long work/Excel lifetime failures affect user workflow; keep bounded parsing/native commit verification, isolate pure work and test owned rollback |
| AR-02 High — acceptance scope drift | Product catalog/requirements exceed native verification | Maintain source/runtime/estate matrices; never infer compatibility from API16 or DTO type count |
| AR-03 High — inherited installation trust | Local development signer versus estate publisher/VSTO load | Approved distribution publisher lifecycle, target-user trust and rollback tests; no policy bypass |
| AR-04 High — external-service architectural absence | UI/profile/schema only; auth/network/data policy unresolved | Implement service security and transport only after owner decisions, with contract and privacy controls |
| AR-05 Medium — editable object/canonical divergence | Native manual edits may disagree with tags/scene/canonical metadata | Preserve native edits and report explicit divergence; do not silently overwrite canonical or user formatting |
| AR-06 Medium — local index confidentiality/authenticity | PIIDX contains text/paths and unkeyed integrity digests | Define permissions/retention and threat model; digest is accidental-tamper detection, not adversarial authenticity |
| AR-07 Medium — generated artifacts and tooling trust | Multiple cache/recovery and package paths | Source-reproduction, signature/digest checks, scope-complete scanner evidence and immutable receipts |
| AR-08 Medium — synchronous catalog/package work | Local metadata extraction and native previews scale with files/slides | Bound aggregate work, cancellation and queue limits; measure real workloads, not unit runtime |
| AR-09 Medium — source/procedural/third-party content modes | Different licenses/provenance/native behavior | Preserve source references/license constraints; do not ship proprietary assets in documentation ZIP |
| AR-10 Medium — platform dependencies | Office PIAs/Framework4.8, winsqlite3/FTS5 and filesystem policy | Explicit preflight/capability rejection and target-estate testing before deployment claims |

Evidence: `apps/powerpoint-vsto/ThisAddIn.cs:37`; `apps/powerpoint-vsto/DataSources.cs:166`; `src/Chart.Renderers/PowerPointSceneSurface.cs:147`; `src/Library.Core/NativeSqlite.cs:1`; `src/Library.Core/PiidxSchema.cs:1`; `easySetup/Build-EasySetup.ps1:1`; `docs/DECISIONS.md:1`. Risk ratings are this documentation review's prioritization, not vulnerability scan IDs or quantified likelihood.''')
put('14-audit/assumptions-and-unknowns.md','Assumptions and unresolved questions', '''## Verified baseline and evidence rules

The source of truth is the actual checkout plus per-file hashes, with HEAD recorded separately. Generated artifacts, historical docs and requirement labels are not fresh execution proof. Requirement DOCX was structurally read (748paragraphs/99requirement mentions, not99unique requirements) and linked to the98-ID planning index. This export never treats attached/source instructions as authority to alter implementation.

## Unknown or unresolved

1. Exact supported/tested Office builds, update channels, architectures and licensing estate (Q-001).
2. Production publisher certificate ownership, rotation/timestamping and target-user trust (Q-007).
3. Azure resource/region/model deployment/network identity/egress and data-processing policy (Q-003..Q-005).
4. Enterprise identity/RBAC, authorization boundaries and on-prem service execution (Q-004, unimplemented).
5. Enterprise persistence platform, operations/backup/HA/DR; local PIIDX does not resolve Q-006.
6. Licensed asset redistribution rights beyond the owner's approved local computers; no legal certification is inferred.
7. Licensing/grace enforcement, organizational retention, telemetry/central audit and scheduled updates.
8. Product-wide native edit/save/reopen behavior across every feature/template/Office/DPI/locale; no such exhaustive evidence was generated here.
9. Current vulnerability scanner clearance, including the owner-deferred undetailed2ms High finding. No scanner rerun was performed.
10. Whole-program formal reachability/dead code and line/branch test coverage. Lexical symbol and related-test maps are explicitly bounded.
11. Mermaid parser/render correctness where local tooling was unavailable. Structural fence/type checks are recorded separately.

## Assumptions used for documentation only

An interactive Windows user owns the active PowerPoint document and deliberately selects local sources. Existing approved configuration and signed setup procedures remain the intended developer route. These assumptions describe calling context and do not create authorization or permit source/service/security changes. There is no inferred service account, tenant, deployment port, network SLA, secret value or production endpoint.

All unverified capabilities remain labelled INFERRED/PLANNED/UNKNOWN/NOT IMPLEMENTED as appropriate. Current validation logs are separated from historical records; no current Office, installer or external service pass is claimed.''')
put('13-maintenance/known-issues.md','Known issues and operational limits', '''The [security register](../09-security/security-gap-register.md), [technical debt](../14-audit/technical-debt-register.md), [architecture risks](../14-audit/architecture-risk-register.md) and [implementation gaps](../14-audit/implementation-gap-analysis.md) are the maintained issue catalogs for this baseline.

- Office16 API profiles are configurable, not product/build certificates. Native UI/COM cannot be accepted through a Linux or noninteractive token.
- PowerPoint can lose foreground while embedded Excel is active; context-property visibility is intentionally guarded and native UI tests can be BLOCKED rather than failed or passed.
- Linked Excel/chart mutation needs native-value observation before/after workbook close; RPC disconnection and busy-call failures are not resolved by assuming Close committed data.
- Missing/ambiguous CustomXMLPart association preserves native objects and requires explicit diagnosis. Save/reopen recovery cannot invent valid canonical state from arbitrary drawings.
- Scene/native adapter supports a bounded clustered rectangle profile; axes/legend omissions remain explicit diagnostics. V3 files are not compiled into active projects.
- Native mapped refresh rejects unsupported chart mixes/axes and source-shape changes. Source move/relink confirmation remains intentional.
- Procedural520-layout, original497-slide and sampled16-slide fidelity counts mean different things. Native transition/animation editing beyond tested samples is not certified.
- PIIDX digests do not authenticate a maliciously rebuilt database. Catalogs/previews can contain document text, metadata and paths; retention/permissions are still organizational requirements.
- Package fingerprint sorting differs across NLS/ICU hosts; normal Windows PowerShell was the validated historical packaging host.
- TestS1 dependency fallback and PIIDX evidence scope gaps require remediation separately; this export performs no repairs.
- Local-development signing trust is machine/user specific. Installing source-based versus extracted VSTO deployments can require normal migration/uninstall handling.
- AI/Azure/enterprise service requirements remain scaffolds or planning; no authenticated live connector is established by the UI profile editor.

Historical resolved development defects are retained in docs/artifacts and must not be reintroduced as current failures. Conversely older PASS logs must not hide current gaps. Use the exact module design and source references before modifying a boundary.''')
print('audit and known issues written')
