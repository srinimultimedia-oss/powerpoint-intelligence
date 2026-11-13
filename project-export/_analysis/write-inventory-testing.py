from pathlib import Path
import json,csv,collections,re,xml.etree.ElementTree as ET,zipfile
R=Path(__file__).resolve().parents[2];E=R/'project-export';inv=list(csv.DictReader((E/'01-repository-analysis/complete-file-inventory.csv').open(encoding='utf-8')));dis=json.loads((E/'_analysis/discovery.json').read_text());tests=json.loads((E/'12-testing/test-structure.json').read_text());runs=json.loads((E/'12-testing/current-validation/results.json').read_text());
def put(path,title,s):(E/path).write_text('# '+title+'\n\n'+s.strip()+'\n',encoding='utf-8')
def table(header,rows):return '| '+' | '.join(header)+' |\n|'+ '|'.join('---' for _ in header)+'|\n'+'\n'.join('| '+' | '.join(str(x).replace('|','\\|').replace('\n',' ') for x in row)+' |' for row in rows)
put('01-repository-analysis/repository-overview.md','Repository overview',f'''**VERIFIED:** baseline commit `ffaec352cbdca666f69f3c402d334108e4115933` (3-Nov-2026), inspected live checkout `C:/Srini/powerpoint-intelligence/powerpoint-intelligence` on13-Nov-2026. The inventory excludes project-export to avoid recursively documenting itself. Every discovered regular file is listed; links are classified without traversal. Git/compiled/vendor/artifact data is counted but not confused with authored implementations.

- Repository inventory: **{dis['files']:,} file/link records**, **{dis['bytes']:,} bytes**.
- Authored code: **{dis['authored_code']}** `.cs/.ps1/.py/.ts/.mjs/.tsx` files outside generated/vendor/VCS/export areas.
- Relevant decoded source/config/schema/documentation text: **{dis['text_files']}** files.
- Symbol locator: **{dis['symbols']:,} lexical candidates**, not an AST-derived API count.
- Inventory read/hash errors: **{len(dis['errors'])}**.

| Area | Actual responsibility | Evidence interpretation |
|---|---|---|
| apps/powerpoint-vsto | VSTO Framework4.8 add-in, WinForms UI, COM features, codecs, local gallery | Active project compile items determine production inclusion |
| src/Chart.Domain | v2 canonical JSON, S0 import; authored V3 files | V3 excluded from active projects |
| src/Chart.Renderers and Shared.Contracts | strict scene contracts/hash, clustered native feasibility renderer, bounded JSON | Native capabilities narrower than generic scene DTO types |
| src/Library.Core | independent PIIDX SQLite/catalog, source extraction, preview/config/watch | Local searchable catalog, not proprietary index loading |
| packages/chart-core/domain/engine | deterministicJS/TS, schema validation and offline scene/SVG tools | Not an OfficeJS frontend or deployed network worker |
| services/ai-gateway | JSON schema and explicit unimplemented README | Planned service only |
| scripts and easySetup | compilation/test/cache/security/diagnostics and dual-bitness signed setup | Some paths write artifacts/install state and are not run during this export |
| tests/.github | custom Console, Node and unittest suites and limited CI | Test existence is not execution/line coverage |
| docs | requirements, decisions, prior evidence and remediation records | Historical/design records are secondary to source |
| artifacts | cached runners, receipts, snapshots, prior evidence/build products | Generated and historical; not additional source modules |

`git-baseline.json` distinguishes235 byte differences from HEAD:234 are CRLF-only; the sole remaining tracked difference is generated `easySetup/README.md`. Source is therefore pinned by actual per-file SHA256 as well as commit, without implying functional uncommitted changes. Untracked ignored deployment assets/dependencies remain in the inventory. No existing export was present, so no prior export was overwritten.''')
stack=[['C#','7.3; .NET Framework4.8','VSTO and standalone domain projects, explicit compile lists','apps/powerpoint-vsto/PowerPointIntelligence.csproj:9'],['Office','PowerPoint COM and VSTO; Office15 PIAs embedded','Desktop Office16 API baseline policy; profiles do not certify2016/2019/365 builds','apps/powerpoint-vsto/OfficeCompatibility.cs:1'],['UI','Windows Forms/System.Drawing','Taskpane, dialogs, floating properties, procedural icons','apps/powerpoint-vsto/ThisAddIn.cs:88'],['TypeScript/JS','Node engine>=24; chart-core/domain>=20','Pure deterministic calculations/strict schemas and offlineCLI; current Nodev24.14.0','packages/chart-engine/package.json:1'],['SQLite','OS winsqlite3.dll; FTS5 capability required','Native bridge and exact PIIDX schema checks; not shippedsqlitebundle','src/Library.Core/NativeSqlite.cs:1'],['JSON','Newtonsoft.Json13.0.3; AJV8.20.0','Runtime domain decoding vs Node schema developer validation','src/Chart.Domain/Chart.Domain.csproj:12'],['Build','VisualStudio/MSBuild/Roslyn, PowerShell','WindowsOffice tools and signing trust; no solutionfile','scripts/Install-Local.ps1:1'],['AI/API','Schema/profile scaffold only','No HTTP server, SDK client, liveAzure calls or identity middleware','services/ai-gateway/README.md:1']]
put('01-repository-analysis/technology-stack.md','Verified technology stack',table(['Technology','Version/target','Role and limitation','Source'],stack)+'\n\nNo React/Electron/browser frontend, Office.js add-in, SmartArt authoring engine, cloud function, ASP.NET runtime service or enterprise database implementation was found in production source. Their presence in requirements does not establish implementation. Dependency currency/CVE status was not queried online; versions below are the inspected lock/project declarations.')
deps=[]
for p in [R/'apps/powerpoint-vsto/PowerPointIntelligence.csproj',R/'src/Chart.Domain/Chart.Domain.csproj']:
 root=ET.fromstring(p.read_text(encoding='utf-8-sig'))
 for n in root.iter():
  tag=n.tag.rsplit('}',1)[-1]
  if tag in ['Reference','PackageReference']:
   deps.append([p.relative_to(R).as_posix(),tag,n.attrib.get('Include',''),n.attrib.get('Version','project/reference identity'),next((a.text or '' for a in n if a.tag.endswith('HintPath')),'framework/tooling resolution')])
for p in sorted((R/'packages').glob('*/package-lock.json')):
 j=json.loads(p.read_text(encoding='utf-8-sig'))
 for name,v in j.get('packages',{}).items():
  if name:deps.append([p.relative_to(R).as_posix(),'locked npm',name,v.get('version',''), 'integrity recorded' if v.get('integrity') else 'no integrity field'])
put('01-repository-analysis/dependency-inventory.md','Dependency inventory',table(['Manifest','Kind','Name','Version','Resolution/integrity'],deps)+'''\n\nInstalled dependency trees and DLLs are separately inventoried as binary/vendor material, not source-authored features. Office/VSTO and winsqlite3 are platform/runtime prerequisites. CI uses actions/checkout@v4 and actions/setup-node@v4; immutable action commit pinning is not present. Lockfiles pin npm versions/integrity but do not authenticate every NuGet/Office/platform dependency. Test-S1 has an independent NuGet download path requiring review (security gap SEC-01), so a lockfile elsewhere must not be generalized to all consumers. No dependencies were updated/installed during this export.''')
entries=[['VSTO host generation','ThisAddIn.Designer generated host item','apps/powerpoint-vsto/ThisAddIn.Designer.cs:1','Office loads signed payload; not a console Main'],['Startup and shutdown','OnStartup/OnShutdown','apps/powerpoint-vsto/ThisAddIn.cs:37','Registers UI/host events and disposes owned resources'],['Ribbon extension','CreateRibbonExtensibilityObject/IntelligenceRibbon','apps/powerpoint-vsto/ThisAddIn.cs:326','XML callback entry, returns DemoRibbon.Create'],['COM automation','RequestComAddInAutomationService/IDemoAutomation','apps/powerpoint-vsto/DemoApi.cs:12','Local IDispatch object, thread-checked mutation'],['Standalone CLI','renderChartV2 main-module guard','packages/chart-engine/tools/render-chart-v2.mjs:1','Reads v2 JSON; writes geometry-only scene to explicit output'],['Installer','Bootstrapper entry and Build-EasySetup','easySetup/Bootstrapper.cs:1','Normal VSTO publisher flow, dualarchitecture selfcontained EXE'],['Local developer install','Install-Local.ps1','scripts/Install-Local.ps1:1','Signs/builds/installs; intentionally not invoked in docs run'],['Test orchestration','Test-All.ps1','scripts/Test-All.ps1:1','Snapshot/cached source proof, local checks and optinOffice'],['Library refresh','LibraryRefreshMonitor','src/Library.Core/LibraryRefreshMonitor.cs:1','Local file watch/timed coordinator; not an enterprise scheduler']]
put('01-repository-analysis/entry-points.md','Application and tooling entry points',table(['Entry','Method/route','Source','Boundary'],entries)+'\n\nNo server Main, controller route, listener or deployed Azure gateway executable is present. Node and Console test Main methods are validation harnesses, not product network interfaces. The [symbol inventory](symbol-inventory.csv) includes declaration candidates for smaller internal/test modules and marks the lexical extraction limitation.')
# Every test file gets an explicit inventory row; assertion-label extraction is a locator, not proof of direct feature coverage.
test_rows=[]
for t in tests:
 path=t['path'];scope='native Office integration' if 'Office' in Path(path).name or 'HostIntegration' in path else 'host-independent/fixture; inspect entry mode'
 test_rows.append([path,t['lines'],', '.join(t['classes']),scope,len(t['assertions']),'; '.join(t['write_indicators']) or 'none found lexically'])
put('12-testing/test-inventory.md','Complete authored test inventory',table(['Source','Lines','Classes','Mode','Literal labels','Write indicators'],test_rows)+'''\n\nThis64-file table covers authored tests/ sources only; seven Node package `.test.mjs` files and PowerShell test wrappers are cataloged in the code inventory and execution tables. Labels are an incomplete lexical locator: loop-generated cases, Require(condition,name) and dynamic descriptions are not counted as distinct execution cases. Exact current Node totals come from the test runner output, not this extraction. `test-structure.json` preserves named assertions/line references and IO indicators for investigation.''')
put('12-testing/test-architecture.md','Testing architecture', '''## Layers and test seams

The repository uses custom C# Console assertions, Node's built-in `node:test` and Python `unittest`; no centralized instrumented coverage tool is configured. Host-free tests exercise codecs, validation, calculations, scene hashing, rollback using fake surfaces, library extraction/schema/cache policies and formatting planning. UI/COM tests are opt-in and require an interactive desktop with valid Office/VSTO activation.

`tests/s3/NativeRendererTests.cs:1` uses a fake `IEditableSceneSurface` with failure injection before/after mutation and checks owned rollback, reentrancy and owner thread. Its interactive mode separately checks real grouping, save/reopen, manual divergence and owned failure cleanup. A fake-surface pass does not prove Office native acceptance.

`tests/s0/HostIntegrationTests.cs:1` exercises embedded workbook edits, bounded native observation, metadata identity, source formatting, duplicate identity and source/destination copy. `ContextPropertiesOfficeTests` observes the installed COM API and distinguishes unavailable foreground focus from a panel mutation failure. No Windows focus policy is changed by those prerequisites.

Library tests have synthetic ZIP/PNG and schema corruption fixtures plus optional Office fidelity runs. Synthetic preview-cache acceptance does not establish native rendering; source-original fidelity combines bounded raster comparison with exact geometry/formatting and selected edit/save/reopen evidence. External licensed corpus fixtures are environmental inputs, not distributed test assets.

## Orchestration

`scripts/Test-All.ps1:1` creates a source snapshot and output under artifacts, preserves stable runner caches via RunnerRoot, runs host-independent checks, architecture builds and optional installed/interactive tests. It validates the snapshot against source fingerprints and uses owned-process timeouts. This export did **not** invoke it, because its filesystem writes violate the export-only assignment. Existing test scripts often compile in artifacts or use normal temp folders; they are documented, not silently redirected with changed application behavior.

CI: four active workflows run portable calculations, Office compatibility configuration, S0 metadata and S1 contracts/schema checks. CI does not load PowerPoint/VSTO or prove enterprise deployment. Root chart-core.yml is a duplicate design file outside .github/workflows, not an additional GitHub Actions workflow.

## Current constrained validation

The export invokes unchanged read-only Node test modules, TypeScript --noEmit, script AST parsing and a direct C# metadata compilation with output and temp under project-export. No existing binaries or receipts are replaced. Logs and exact Node argument arrays/times are stored beside this documentation. Source-preservation verification provides a separate gate from test success.''')
rt=[]
for x in runs:rt.append([x['name'],x['status'],x['counts'].get('tests','typecheck'),x['counts'].get('pass','n/a'),x['start_utc'],x['end_utc'],x['log']])
put('12-testing/test-results.md','Current validation and historical evidence',table(['Current check','Result','Cases','Passes','Start UTC','End UTC','Evidence'],rt)+'''\n\nAdditional **current** checks:17 S0 metadata assertions passed from unchanged source, compiled/run entirely in export-local temporary storage; see [metadata log](current-validation/metadata-tests.log) and `_analysis/run-metadata.ps1` in the full local export for the exact command. Fifty PowerShell files parsed without syntax errors; see [AST log](current-validation/powershell-ast.log) and [per-file results](current-validation/powershell-ast.json). Parsing executes none of those scripts. Metadata and AST evidence files' filesystem timestamps identify this export run; Node JSON records precise start/end UTC. Node runtime v24.14.0.

Total currently executed named assertions: **226 =6 calculations +25 domain/schema +178 engine +17 metadata**. Type-checking and50-file syntax parsing are separate successful checks, not226 more tests. No line/branch coverage, mutation testing or independent production packaging build was run.

**NOT EXECUTED now:** full Test-All, Office automation, installed ribbon, linked Excel native refresh, original template fidelity, renderer native save/reopen, x86/x64 host process integration, installer signing/trust/install/rollback, external service calls, scanner rescan, PIIDX corpus execution and security scripts that write outside export. None is a pass inferred from absence of a failure.

**Historical only:** docs/implementation/deployment-validation-20261103.md records24/24 local checks and desktop package1.0.0.18 validation. Existing artifacts contain older results, including failures during development. Such source records establish history; they are not13-Nov current execution evidence. Earlier Office successful focused assertions do not cover every chart family, every template, DPI/locale or Office estate. Historical Checkmarx remediation is not a fresh vulnerability scan.''')
put('12-testing/coverage-gap-analysis.md','Test coverage gaps', '''There is no collected line/branch coverage denominator in the repository. Source/test association is traced in test-feature-traceability.csv, but an existing related test path is not proof every command branch or API was exercised. The current constrained run intentionally covers deterministic source/metadata behavior only.

| Gap | Evidence and risk | Required validation |
|---|---|---|
| Office estate | Profiles share API16; product/build/architecture matrix unproven by configuration tests | Interactive2016/2019/365 builds andx86/x64 package loading, not just compiler targets |
| Native renderer breadth | Generic scene DTOs/TS support more geometries than native clustered rectangle adapter | Per-family capability tests and deliberate rejection before mutation |
| V3 production | V3 source omitted compile lists; no direct codec harness found | Explicit build/migration decision and direct boundary tests before wiring |
| Full feature command behavior | Catalog/ribbon shape checks exceed native acceptance coverage | Every routed command, unsupported selections, undo, cancellation and persistence |
| Gallery/native templates | Procedural layouts and original source imports are different modes;16-original historical sample | Full intended native corpus/formatting edit/save/reopen and license-aware asset tests |
| DPI/foreground/locale | Context controls depend on native focus, selection and WinForms layout | Multi-monitor scaling, display transitions, keyboard/accessibility and locale matrix |
| Data link refresh | Primary-axis uniform category charts only in mapped refresh | Moved/header/shape changes, XY/pie/secondary axes, long-lived Excel lifetime if expanded |
| AI/enterprise | No implementation, so labels/schema cannot be end-to-end tests | Authenticated transport/governance/approval/audit tests only after implementation |
| Build supply chain | TestS1 fallback download lacks independent digest; PIIDX scanner evidence filter mismatch | Tampered package tests and positive detection fixtures covering actual runner paths |
| Host-globalization fingerprint | NLS/ICU ordering can yielddifferentpackaging fingerprints | Cross-PS5/PS7 ordinal canonical vectors; no recipechange made here |

This export's completeness metric measures documentation/traceability and bounded validation, not product acceptance. Planned/missing modules remain visible rather than being promoted to verified implementations.''')
print('inventory and testing documents written')
