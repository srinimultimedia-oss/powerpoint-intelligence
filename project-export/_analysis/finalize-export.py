from pathlib import Path
import json,csv,re,os,hashlib,datetime,subprocess,zipfile,collections
R=Path(__file__).resolve().parents[2];E=R/'project-export';MASTER=E/'15-deliverables/MASTER-TECHNICAL-DESIGN.md';NOW=datetime.datetime.now(datetime.timezone.utc).isoformat();COMMIT='ffaec352cbdca666f69f3c402d334108e4115933'
required={
'':['README.md','EXPORT-MANIFEST.md','DOCUMENTATION-INDEX.md','EXECUTIVE-TECHNICAL-SUMMARY.md'],
'01-repository-analysis':['repository-overview.md','complete-file-inventory.csv','technology-stack.md','dependency-inventory.md','entry-points.md','source-analysis-coverage.md'],
'02-architecture':['enterprise-architecture.md','solution-architecture.md','application-architecture.md','component-architecture.md','deployment-architecture.md','architecture-decisions.md','architecture-diagrams.md'],
'03-modules':['module-inventory.md','module-dependency-matrix.md'],
'04-powerpoint-integration':['powerpoint-integration-overview.md','addin-lifecycle.md','ribbon-architecture.md','native-editing-engine.md','shape-and-slide-operations.md','rendering-engine.md','interoperability.md'],
'05-features':['complete-feature-inventory.md','feature-implementation-details.md','feature-source-mapping.csv','feature-status-matrix.md','feature-workflows.md'],
'06-ai-architecture':['ai-solution-architecture.md','azure-ai-integration.md','prompt-orchestration.md','ai-workflows.md','ai-data-flow.md','ai-security.md'],
'07-api-design':['api-inventory.md','internal-interfaces.md','external-interfaces.md','request-response-contracts.md'],
'08-data-design':['data-models.md','class-diagrams.md','data-dictionary.md','persistence-design.md','serialization-contracts.md'],
'09-security':['security-architecture.md','authentication-authorization.md','secrets-management.md','threat-model.md','security-gap-register.md'],
'10-workflows':['workflow-inventory.md','sequence-diagrams.md','activity-diagrams.md','error-handling-flows.md','data-flow-diagrams.md'],
'11-build-deployment':['development-setup.md','build-instructions.md','packaging-installation.md','deployment-guide.md','configuration-reference.md','troubleshooting-guide.md'],
'12-testing':['test-architecture.md','test-inventory.md','test-results.md','test-feature-traceability.csv','coverage-gap-analysis.md'],
'13-maintenance':['developer-onboarding.md','debugging-guide.md','extension-guide.md','operational-runbook.md','known-issues.md'],
'14-audit':['technical-debt-register.md','implementation-gap-analysis.md','architecture-risk-register.md','evidence-traceability.csv','documentation-validation.md','assumptions-and-unknowns.md'],
'15-deliverables':['MASTER-TECHNICAL-DESIGN.md','TECHNICAL-HANDOVER.md','DOCUMENTATION-COMPLETENESS-REPORT.md']}
required_paths=[str(Path(d)/n).replace('\\','/') for d,ns in required.items() for n in ns]
(E/'_analysis/required-files.json').write_text(json.dumps(required_paths,indent=2),encoding='utf-8')
modules=json.loads((E/'03-modules/module-map.json').read_text());features=list(csv.DictReader((E/'05-features/feature-source-mapping.csv').open(encoding='utf-8-sig')))
analysis=json.loads((E/'14-audit/analysis-coverage.json').read_text());interfaces=list(csv.DictReader((E/'07-api-design/interface-coverage.csv').open(encoding='utf-8-sig')))
def put(p,title,s):(E/p).write_text('# '+title+'\n\n'+s.strip()+'\n',encoding='utf-8')
def slug(s):return re.sub(r'[^\w\- ]','',s.replace('`','').lower()).replace(' ','-')
def toc(text):
 if len(text.splitlines())<100 or '<!-- generated-toc -->' in text:return text
 hs=[m.group(1) for m in re.finditer(r'^## (.+)$',text,re.M)]
 if len(hs)<3:return text
 lines=text.splitlines();insert='<!-- generated-toc -->\n## Contents\n\n'+'\n'.join('- ['+x+'](#'+slug(x)+')' for x in hs)+'\n<!-- end-generated-toc -->\n'
 return lines[0]+'\n\n'+insert+'\n'+'\n'.join(lines[1:]).lstrip()+'\n'
# Strip residual inline bodies from operations declaration metadata, preserving explanatory prose.
p=E/'09-security/reviewed-files.csv';rows=list(csv.DictReader(p.open(encoding='utf-8-sig')))
if rows and 'functions' in rows[0]:
 fields=[x for x in rows[0] if x!='functions']
 with p.open('w',encoding='utf-8',newline='') as f:
  w=csv.DictWriter(f,fieldnames=fields,extrasaction='ignore');w.writeheader();w.writerows(rows)
for p in (E/'03-modules/individual-module-designs/operations').glob('*.md'):
 text=p.read_text(encoding='utf-8');ls=[]
 for line in text.splitlines():
  if line.startswith('| `') and ' | Complete text read;' in line:
   cells=line.split('|');names=re.findall(r'(?:\d+:)?(?:public|internal|private|protected)\s+(?:(?:static|sealed|override|virtual|async)\s+)*[\w<>\[\].,?]+\s+(\w+)\s*\(|\bfunction\s+([\w-]+)',cells[-2]);nm=[next(v for v in t if v) for t in names];cells[-2]=' '+('Declarations: '+', '.join(dict.fromkeys(nm)) if nm else 'Script/project entry; its actual flow and responsibilities are documented above.')+' ';line='|'.join(cells)
  ls.append(line)
 p.write_text('\n'.join(ls)+'\n',encoding='utf-8')
# Validation log with no stdout is real, not a placeholder.
p=E/'12-testing/current-validation/typescript-noemit.log'
if p.exists() and not p.read_bytes():p.write_text('Wrapper note: compiler exited0 and produced no stdout/stderr. Exact command/times are recorded in results.json.\n',encoding='utf-8')
put('README.md','NexWave technical design export',f'''
Documentation export version1.0, generated13-Nov-2026 (Asia/Dubai); baseline commit `{COMMIT}`. All application files were preserved; only this export folder was written. Start with the [standalone master](15-deliverables/MASTER-TECHNICAL-DESIGN.md), [documentation index](DOCUMENTATION-INDEX.md), [technical handover](15-deliverables/TECHNICAL-HANDOVER.md) or [completeness report](15-deliverables/DOCUMENTATION-COMPLETENESS-REPORT.md).

The inspected product is a Windows Framework4.8 VSTO/COM add-in with WinForms controls, local linked-data/native-editing tools and an independent PIIDX SQLite catalog. Node/TypeScript calculations and scene tooling are offline developer boundaries. Azure/on-prem services are schema/profile scaffolds, not implemented production clients/servers.

This package documents46coherent modules and190feature entries (172UI command routes+18cross-cutting capabilities). Every204authored code file is accounted for in source-specific review manifests;46VSTO code files include targeted-body/full-text indexing rather than formal line-by-line proof. Supporting/generated/vendor/binary/private material is explicitly classified in the30,301record inventory.

Evidence labels: **VERIFIED**=inspected source or explicitly executed check; **INFERRED**=supported deduction; **PLANNED**=roadmap/specification; **UNKNOWN**=insufficient evidence; **NOT IMPLEMENTED**=absent from inspected execution graph. A source-implemented operation is not a native acceptance pass.

Current safe validation:229assertions pass (209Node+17metadata+3read-only path controls), TypeScript noEmit and50PowerShell AST parses pass. No current Office/installer/scanner/external-service execution is claimed. Historical logs remain historical. Mermaid syntax/render parsing was unavailable; structural checks and original diagram counts are reported separately.

All source citations are repository-relative `path:line`, resolved against the baseline checkout, not copied source files. Internal document links remain portable. The [ZIP](15-deliverables/PowerPoint-Intelligence-NexWave-Technical-Design.zip) excludes authoring/validation scripts, temporary binaries, application source, licensed assets and private configuration. The complete local export retains audit machinery under_analysis/.unlazy; those are not product components or portable deliverables. Checksums and ZIP verification are recorded in the manifest/audit outputs.

No API OpenAPI spec is supplied because no actual HTTP endpoint implementation exists; [applicability](07-api-design/openapi/README.md) explains the deliberate omission. This export is not a claim of full feature parity, production certification or zero security findings.''')
put('EXECUTIVE-TECHNICAL-SUMMARY.md','Executive technical summary','''## Actual system

NexWave is a Windows desktop PowerPoint add-in, hosted inside Office through VSTO/COM and .NETFramework4.8. WinForms provides taskpane/dialog/context controls. Native charts/tables/shapes remain PowerPoint-editable, with versioned CustomXMLPart/tag associations. Local Excel/CSV/JSON/XML links and explicit worksheet/range mappings feed chart/table operations. A local PIIDX SQLite/FTS catalog indexes PPTX/POTX metadata and native preview/source-import workflows. PureJS/TS calculations and strict scene contracts are separate developer tooling.

## Implementation distinctions

V3 chart source exists but is not in active compile lists. Generic scene DTOs and the six-family TS engine do not establish native renderer coverage: the semantic native surface is a bounded clustered rectangle profile with explicit omissions. Source-original497slides, procedural520layouts and16sampled original fidelity cases measure different things. Managed elements and contextual controls have explicit selection/metadata/geometry limits.

No authenticated Azure client, running on-prem service/API, enterprise RBAC/SSO, governed model-action executor or central scheduler is implemented. Profiles/schema and approved architecture intentions remain scaffolds/planning.

## Engineering priorities

The audit identifies an independently unverified NuGet fallback download in TestS1; scanner-evidence scope excludingPIIDX; related implicitXML readers; source/procedural/proprietary-content provenance distinctions; native acceptance/DPI/Office-estate gaps; host-dependent package fingerprints; and test-oracle/cleanup weaknesses. These are documented, not fixed, in this assignment.

Current constrained validation passes229assertions, TypeScript type checking and50PowerShell syntax parses. No new Office integration, signed deployment or vulnerability scan was run. Complete source-preservation hashes, source/test/module maps and explicit metric denominators support engineering handover without implying complete product acceptance.''')
put('15-deliverables/TECHNICAL-HANDOVER.md','Technical handover','''## First orientation

Read the [master design](MASTER-TECHNICAL-DESIGN.md), [module inventory](../03-modules/module-inventory.md), [actual feature status](../05-features/feature-status-matrix.md), [current test evidence](../12-testing/test-results.md) and [security gaps](../09-security/security-gap-register.md). Match commit and per-file hashes before using line references. Visible brand is NexWave; existing PowerPointIntelligence COM/metadata/storage identities remain compatibility contracts.

## Runtime maintenance map

ThisAddIn owns startup, callbacks, UI gating and cleanup. PowerPointHost owns S0 chart/metadata association and UI STA checks. DataSources and mapped-link helpers own bounded input/provenance/commit observation. Element modules own text/process/connector/indicator/agenda state. RendererRibbonCommand/native surface owns the narrower grouped-shape projection. Gallery/workspace/Library.Core own local registration/extraction/search/preview/import. Module maps give all source assignments and method locators.

## Build and change discipline

Use VisualStudio Office tools/Framework4.8 and the approved signer in normal Windows PowerShell. Test deterministic contracts first, then explicit native editing/save-reopen, and finally installed package/ribbon acceptance. Changing DTO definitions outside the compile graph is not a feature release; changing persisted namespace/schema requires explicit migration. Avoid broad clean/delete operations that remove installed extraction paths, stable receipts or security evidence.

## Immediate backlog

Address source-backed supply-chain/scanner-scope gaps and inconsistentXML policy; strengthen vacuous/structural test oracles and owned Office cleanup; define distribution publisher/Office matrix/data retention; make fingerprint ordering host-independent in a separately tested signed release. EnterpriseAI/auth/services need decisions and implementation before UI labels can promise live behavior. Do not import undocumented native/proprietary behavior or distribute licensed source assets beyond approved scope.

## Validation boundary

This export changes no application code and performs no install or dependency upgrade. Safe current checks cover229assertions, TS typecheck and50script parses. Office, live services, freshscanner clearance and production compatibility remain unverified here. Full-body/targeted review levels, exclusion reasons and diagram parser limitations are explicit. Use the issue registers as an implementation backlog, not completed remediation evidence.''')
# TOC insertion for long supporting docs before assembly.
for p in E.rglob('*.md'):
 if '_analysis' in p.parts or '.unlazy' in p.parts or p==MASTER:continue
 p.write_text(toc(p.read_text(encoding='utf-8-sig')),encoding='utf-8')
def include(name):
 p=E/name;t=p.read_text(encoding='utf-8-sig');t=re.sub(r'<!-- generated-toc -->.*?<!-- end-generated-toc -->\s*','',t,flags=re.S);title=t.splitlines()[0].lstrip('# ');t='\n'.join(t.splitlines()[1:]).lstrip()
 def link(m):
  label,target=m.groups()
  if re.match(r'^(https?://|mailto:)',target):return m.group()
  target=target.strip('<>');n,sep,frag=target.partition('#');dest=(p.parent/n).resolve() if n else p
  return '['+label+']('+os.path.relpath(dest,MASTER.parent).replace('\\','/')+(('#'+frag) if sep else '')+')'
 t=re.sub(r'\[([^\]\n]*)\]\(([^\n)]+)\)',link,t)
 t=re.sub(r'^(#{1,6}) ',lambda m:'#'*min(6,len(m.group(1))+2)+' ',t,flags=re.M)
 return '\n### '+title+'\n\n'+t+'\n'
sections=[
('1. Document control and baseline',None),('2. Executive technical overview',['EXECUTIVE-TECHNICAL-SUMMARY.md']),('3. System context and scope',['01-repository-analysis/repository-overview.md']),('4. Verified technology stack and dependencies',['01-repository-analysis/technology-stack.md','01-repository-analysis/dependency-inventory.md']),('5. Architecture principles and decisions',['02-architecture/architecture-decisions.md']),('6. Enterprise and solution architecture',['02-architecture/enterprise-architecture.md','02-architecture/solution-architecture.md']),('7. Application lifecycle and UI architecture',['02-architecture/application-architecture.md']),('8. Component and module architecture',['02-architecture/component-architecture.md','03-modules/module-inventory.md']+[m['document'] for m in modules]),('9. Detailed feature design',['05-features/complete-feature-inventory.md','05-features/feature-implementation-details.md','05-features/feature-status-matrix.md']),('10. PowerPoint integration',['04-powerpoint-integration/powerpoint-integration-overview.md','04-powerpoint-integration/addin-lifecycle.md','04-powerpoint-integration/ribbon-architecture.md','04-powerpoint-integration/interoperability.md']),('11. Rendering and native editing',['04-powerpoint-integration/native-editing-engine.md','04-powerpoint-integration/shape-and-slide-operations.md','04-powerpoint-integration/rendering-engine.md']),('12. AI integration architecture',['06-ai-architecture/ai-solution-architecture.md','06-ai-architecture/azure-ai-integration.md','06-ai-architecture/prompt-orchestration.md','06-ai-architecture/ai-workflows.md','06-ai-architecture/ai-data-flow.md','06-ai-architecture/ai-security.md']),('13. API and interface specifications',['07-api-design/api-inventory.md','07-api-design/internal-interfaces.md','07-api-design/external-interfaces.md','07-api-design/request-response-contracts.md']),('14. Data models and persistence',['08-data-design/data-models.md','08-data-design/class-diagrams.md','08-data-design/data-dictionary.md','08-data-design/persistence-design.md','08-data-design/serialization-contracts.md']),('15. Security architecture',['09-security/security-architecture.md','09-security/authentication-authorization.md','09-security/secrets-management.md','09-security/threat-model.md','09-security/security-gap-register.md']),('16. Runtime and deployment architecture',['02-architecture/deployment-architecture.md']),('17. Build installation configuration and release',['11-build-deployment/development-setup.md','11-build-deployment/build-instructions.md','11-build-deployment/packaging-installation.md','11-build-deployment/deployment-guide.md','11-build-deployment/configuration-reference.md','11-build-deployment/troubleshooting-guide.md']),('18. Execution sequences and activity diagrams',['10-workflows/workflow-inventory.md','10-workflows/sequence-diagrams.md','10-workflows/activity-diagrams.md','10-workflows/data-flow-diagrams.md']),('19. Error handling and observability',['10-workflows/error-handling-flows.md']),('20. Testing architecture and results',['12-testing/test-architecture.md','12-testing/test-inventory.md','12-testing/test-results.md','12-testing/coverage-gap-analysis.md']),('21. Performance and scalability',None),('22. Operational maintenance',['13-maintenance/operational-runbook.md','13-maintenance/debugging-guide.md']),('23. Known limitations and technical debt',['14-audit/technical-debt-register.md','14-audit/implementation-gap-analysis.md','14-audit/architecture-risk-register.md','13-maintenance/known-issues.md']),('24. Extension and development guidance',['13-maintenance/developer-onboarding.md','13-maintenance/extension-guide.md']),('25. Implementation and documentation traceability',None),('26. Assumptions unknowns and unresolved questions',['14-audit/assumptions-and-unknowns.md']),('27. Appendices and supporting references',None)]
out=['# NexWave — Master Technical Design','','Standalone evidence-grounded engineering handover. This document includes all46module designs and the complete190entry feature design; supporting CSV inventories preserve exhaustive mappings. Implementation source is not redistributed.','','## Contents','']+['- ['+t+'](#'+slug(t)+')' for t,_ in sections]
for title,docs in sections:
 out+=['','## '+title,'']
 if docs:
  for d in docs:out.append(include(d))
 elif title.startswith('1.'):
  out.append(f'Export1.0; generated13-Nov-2026, Asia/Dubai; UTC {NOW}. Repository baseline commit `{COMMIT}`. Exact byte/source preservation is in14-audit/source-preservation.json;234tracked byte differences are GitEOL normalization only and generated easySetup/README.md is the sole non-EOLtracked difference. Source references use repository-relative path:line. No application code/config/dependency was changed. VERIFIED source/currentexecution, INFERRED, PLANNED, UNKNOWN and NOTIMPLEMENTED are distinct evidence labels. Review caveats are substantive: file accounting is not line-by-line proof or native acceptance.')
 elif title.startswith('21.'):
  out.append('No production throughput/SLA/HA or current native performance benchmark was measured. Local bounded parsers/semantic validation constrain individual inputs; the native clustered adapter caps2000nodes and explicitly rejects unsupported geometry, while default simple Excel loading bounds2–50rows and2–15columns. PIIDX uses transactions/indexed/FTS paging and a background refresh coordinator but native previews and COM updates remain desktop resources. File hashes/integrity validation have IO cost even for unchanged sources. Aggregate package decompression budget and cancellation/stress coverage remain review gaps. Current test durations in12-testing/current-validation/results.json are unit-run times, not application benchmarks. Historical corpus timings are environment-specific and must not be generalized to estate capacity. Windows UI STA work can block interaction; optimize pure work with ownership/cancellation and regression tests rather than performing COM on arbitrary background threads.')
 elif title.startswith('25.'):
  out.append('File inventory: [complete-file-inventory.csv](../01-repository-analysis/complete-file-inventory.csv). Module assignments: [source-module-map.csv](../03-modules/source-module-map.csv). Every190feature entry: [feature-source-mapping.csv](../05-features/feature-source-mapping.csv). All71test sources: [test-feature-traceability.csv](../12-testing/test-feature-traceability.csv). Validated citations: [evidence-traceability.csv](../14-audit/evidence-traceability.csv). Review levels: [source analysis](../01-repository-analysis/source-analysis-coverage.md). Counts are scoped explicitly; command routes are not an assertion of190independent completebusiness features. Full-body test reviews distinguish textual/mocked/candidate mappings from native execution and expose vacuous or shallow assertions.')
 else:out.append('See [documentation index](../DOCUMENTATION-INDEX.md), [export manifest](../EXPORT-MANIFEST.md) and [completeness report](DOCUMENTATION-COMPLETENESS-REPORT.md). OpenAPI is intentionally inapplicable because no HTTP endpoint implementation exists. PortableZIP carries documentation/validationdata only, excluding source, private configuration, dependencies, licensed assets, helper scripts and binaries. Primary source/decision/requirement documents remain in their original repository locations. This export is a reproducible technical documentation baseline, not a production acceptance or scanner-clearance certificate.')
MASTER.write_text('\n'.join(out)+'\n',encoding='utf-8')
print(json.dumps({'required_document_paths':len(required_paths),'master_bytes':MASTER.stat().st_size,'modules':len(modules),'feature_entries':len(features)}))
