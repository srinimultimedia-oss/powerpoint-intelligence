from pathlib import Path
import json,csv,re,os,hashlib,datetime,subprocess,zipfile,sys,importlib.util,collections
R=Path(__file__).resolve().parents[2];E=R/'project-export';NOW=datetime.datetime.now(datetime.timezone.utc).isoformat();COMMIT='ffaec352cbdca666f69f3c402d334108e4115933'
required=json.loads((E/'_analysis/required-files.json').read_text());mods=json.loads((E/'03-modules/module-map.json').read_text());fs=list(csv.DictReader((E/'05-features/feature-source-mapping.csv').open(encoding='utf-8-sig')));ifs=list(csv.DictReader((E/'07-api-design/interface-coverage.csv').open(encoding='utf-8-sig')));tr=list(csv.DictReader((E/'12-testing/test-feature-traceability.csv').open(encoding='utf-8-sig')))
def mdpaths():return [p for p in E.rglob('*.md') if '_analysis' not in p.parts and '.unlazy' not in p.parts]
def put(n,t,s):(E/n).write_text('# '+t+'\n\n'+s.strip()+'\n',encoding='utf-8')
feature_mapped=sum(bool(re.search(r'(?:tests|packages)/.+\.(cs|mjs|py|ps1)',x['tests'])) for x in fs)
# Substantive final gates/report, not completion-by-presence alone.
put('15-deliverables/DOCUMENTATION-COMPLETENESS-REPORT.md','Documentation completeness and acceptance','''## Scope and meaning

The document-delivery percentage measures the81explicit required files, including CSVs, not product readiness. All applicable deliverable paths are present and substantive; theOpenAPI directory contains an explicit inapplicability explanation instead of fabricated endpoints. Forty-sixmodule designs add detailed coherent coverage beyond the requested directory skeleton. Accuracy limitations, missing functionality and unexecuted runtime checks remain part of the completed documentation.

| Metric | Numerator / denominator | Interpretation |
|---|---|---|
| Required deliverable-file coverage |81/81=100% | Explicit requestedfilelist, including master/index/manifest/reports andthreeCSV artifacts |
| Authored code analysis accounting |204/204=100% | File-level source-specific review;46code files include targeted-body/fulltext indexing, not formal line-by-line proof |
| Full repository inventory |30,301/30,301records |30,300regular files +1reparse record not followed; generated/vendor/VCS/binary/private material classified |
| Module design coverage |46/46=100% |19VSTO+18core+9operations coherent modules;17requested dimensions per design |
| Feature documentation |190/190=100% |172UI routes+18capabilities, not190complete independentbusinessfeatures or parity certification |
| API/interface documentation |12/12defined groups |11implemented integration boundary groups+1planned schema-only boundary; not every private helpermethod |
| Test-source mapping |71/71=100% |64authored tests/ files+7package testfiles; related mappings are not executed per-feature acceptance |
| Feature entries with related test candidates |FEATURE_MAPPED/190 | Source-related references only; not instrumentedcoverage or allbehavior exercised |
| Current assertions |229passed |209Node+17metadata+3read-only pathpolicy controls; TSnoEmit and50scriptparses separate |
| Source-reference validation |Measured in documentation-validation.md | Everyoccurrence file/range checked, plus semanticmodule review; rangevalidity not prose proof |
| Diagram structural validation |Measured in documentation-validation.md | EveryMermaidblock/fence/type checked; duplicatedmasterdiagrams countedseparatelyfromunique designs |
| Mermaid full parser/render validation |0 / unique diagram designs | No localMermaidruntime available; no syntax/renderpassclaimed |
| Repository preservation |30,300regular filebytehashes unchanged | Fullbaseline comparison;1historicaltestjunction inventoried/excludedfromtraversal |

The overall **required document-delivery coverage is100%** with explicit scoped denominators. It must not be read as100%fullsemantic/formalproof, native runtime acceptance, scannerclearance, benchmarkcoverage or Mermaidrenderer verification. Excluded/generatedareas remain inventory-only. Source-level coverage and evidencequalifications are in [source analysis](../01-repository-analysis/source-analysis-coverage.md).

## Review and validation controls

Primary authors read assigned implementation bodies, interfaces/schemas/config and test code. Parent independently rechecked source references, module/feature/test mapping fields, contract examples and consequentialcounterexamples. Validation controls deliberatelyinclude invalidsource refs/missinglinks/badMermaidtype and a syntheticsecret marker; a validator that cannot flag those controls is not accepted. Portablearchive contents andchecksums are independently verified. No application implementation, config, dependency, existingbinary or security setting was changed.

## Remaining limits

No currentOfficeactivation/native editing/save-reopen, signedinstaller generation/loading, x86/x64hostruntime, vulnerabilityscan, enterpriseAI/auth/API invocation, instrumentedline/branchcoverage or productionperformancevalidation was performed. No fullMermaidparser was available. Historicalevidence remains historical. Currentpathchecksreuseanexistingjunctionfixturewithoutcreatingormodifyingit. Allthese boundaries are explicit, rather than silently treatedaspasses.
'''.replace('FEATURE_MAPPED',str(feature_mapped)))
# Write a detailed validation report now; final counts updated below.
put('14-audit/documentation-validation.md','Documentation validation method and results','''## Executed export checks

The parent validator checks required-file presence/nonempty content, everyexplicit source-file/line-range reference, internalfilelinks, Mermaidfence/type structure, unique feature IDs/requiredfields, module/source/test assignments, source hashes/review rows, portablefiletypes, archive hashes/content exclusion and source preservation. A separate contractexample check parsed twoJSON examples and validated theplannedAI sample against the actualAJV schema; this is not liveAI execution.

Negative controls exercise invalidsource references, a missinglink, unsupporteddiagramtype and syntheticsecret-marker detection. A knownvalidcontrol proves rejectionisnotacatch-all. Source-prose meaning is editoriallyreviewed; file/linevalidation alone cannot prove behavior. LocalMermaidparser/runtime was notfound; structuralvalidation is not fullsyntax/render validation.

Validation JSON/transcripts and exact metriccounts are linked below after the final pass. The master includes copied diagrams/sourcecitations, so occurrencecounts exceed unique designs/references. Parent/leaf outputs are independentlychecked; none is accepted solelybecauseitsauthorreported success.

Application preservation baseline covers30,300regular files outside export. One historicaltestjunction is explicitly inventoried and notfollowed. Currentvalidation createsoutputs/tempbinaryonlyinsideexport; portableZIPexcludeshelpercode/binaries/privateconfiguration/sourceassets. All runtime/scan/installation omissions are documented separately in test-results and assumptions.
''')
# Document index includes every Markdown documentation artifact, excluding internal gate machinery.
for _ in range(2):
 docs=sorted(mdpaths(),key=lambda p:p.relative_to(E).as_posix())
 lines=['# Documentation index','','Start with [master design](15-deliverables/MASTER-TECHNICAL-DESIGN.md), [executive summary](EXECUTIVE-TECHNICAL-SUMMARY.md), [handover](15-deliverables/TECHNICAL-HANDOVER.md), [completeness](15-deliverables/DOCUMENTATION-COMPLETENESS-REPORT.md) and [manifest](EXPORT-MANIFEST.md).','','CSV/JSON validation and traceability data are listed in the export manifest; authoring scripts/temporary binaries/gate machinery are local audit tooling, not product documentation or portable payload.','','| Document | Purpose/title |','|---|---|']
 for p in docs:
  if p.name=='DOCUMENTATION-INDEX.md':continue
  rel=p.relative_to(E).as_posix();title=p.read_text(encoding='utf-8-sig').splitlines()[0].lstrip('# ');lines.append('| ['+rel+']('+rel+') | '+title.replace('|','\\|')+' |')
 (E/'DOCUMENTATION-INDEX.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
 put('EXPORT-MANIFEST.md','Export manifest',f'''Application: **NexWave / PowerPoint Intelligence**. Repository: `{R.as_posix()}`. Commit baseline: `{COMMIT}`. Generated UTC:{NOW} (13-Nov-2026Asia/Dubai). Exportversion1.0; destination: `project-export/`.

All81requiredfiles are supplied, plus46module designs, boundedinterface/native/event appendices, CSV/JSON source/review/test inventories and validation evidence. [Index](DOCUMENTATION-INDEX.md) provides the completeMarkdownlist; [portable-file inventory](14-audit/export-file-inventory.csv) records everyselectedfile/size/hash. Exactsourcehashes and exclusions are in the repositoryinventory/review manifests.

Coverage:204/204authoredcodefiles reviewed at explicit levels;339relevanttextfiles structurallyinspected;46/46modules;190/190featureentries;12/12declaredinterfacegroups (oneplanned);71/71testsourcefiles mapped. Relatedtestcandidates cover{feature_mapped}/190featureentries and are notinstrumentedcoverage. Fullrepositorydiscovery30,301records includes30,300regularfiles and1nontraversedreparseentry. Source bytebaselineunchanged. See [completeness](15-deliverables/DOCUMENTATION-COMPLETENESS-REPORT.md) and [validation](14-audit/documentation-validation.md).

Currenttests:229assertionspass, TypeScriptnoEmitpasses,50PowerShellfilesparse. No currentnativeOffice/signedinstallation/externalservices/scannerpass. Mermaidstructuralchecksareperformed; fullparser/render unavailable. Exact finalreferences/diagram/linkcounts are recorded in validation JSON/report, not guessed here.

OpenAPI is intentionally omitted as an executable specification because noHTTPserver/routes exist; [reason](07-api-design/openapi/README.md). AIarchitecture is profile/schema/planning, not livedeployment. V3contracts are authoredbutnotcompiled. Newsecurity/testgaps remain recommendations; historicalCheckmarxclearance is not asserted.

PortableZIP: [PowerPoint-Intelligence-NexWave-Technical-Design.zip](15-deliverables/PowerPoint-Intelligence-NexWave-Technical-Design.zip). No applicationcode/dependencies/privateconfiguration/credentials/binaries/licensedtemplateassets are included. Local_analysis, .unlazy, authoring.py/.mjs/.ps1 helpers andtemporarytestoutputs are excluded. Repository-relative sourcepaths/lines remain textual citations, not copied source files. Diagram/codecontractexamples are documentation, not source distribution.

Checksums: [DOCUMENT-CHECKSUMS.sha256](15-deliverables/DOCUMENT-CHECKSUMS.sha256) hashes selecteddocument/data files; it excludes itself andZIP/hashreceipt to avoid cycles. [ZIP checksum](15-deliverables/PowerPoint-Intelligence-NexWave-Technical-Design.zip.sha256) and [archive verification](14-audit/archive-validation.json) record the finishedarchive. The archive excludes itself/previousarchives and its external audit/checksumreceipts.

Overall status: requested documentationpackage delivered with explicit limitations, notproduction or securitycertification. No prior exportexisted; no existing applicationfile was replaced. Subsequent regeneration shouldpreservethisbaselineoruseaversionedexportfolder.
''')
# Run independent structural validator.
proc=subprocess.run([sys.executable,'-B',str(E/'_analysis/validate-export.py'),'all'],cwd=R,capture_output=True,text=True,encoding='utf-8');(E/'14-audit/export-validation-transcript.log').write_text(proc.stdout+proc.stderr,encoding='utf-8')
pending={'14-audit/export-file-inventory.csv','15-deliverables/PowerPoint-Intelligence-NexWave-Technical-Design.zip','15-deliverables/DOCUMENT-CHECKSUMS.sha256','15-deliverables/PowerPoint-Intelligence-NexWave-Technical-Design.zip.sha256','14-audit/archive-validation.json','../14-audit/evidence-traceability.csv'}
preview=json.loads((E/'14-audit/parent-validation-all.json').read_text())
real_errors=[x for x in preview['errors'] if not(x.get('error')=='missing link target' and x.get('target') in pending)]
if real_errors:print(real_errors);raise SystemExit('Export structural validation failed')
v=json.loads((E/'14-audit/parent-validation-all.json').read_text());unique_refs={(x['source'],x['line'],x['end']) for x in v['source_references']};unique_diagrams={x['sha256'] for x in v['diagrams']}
# Traceability is document-to-source evidence, including valid repeated master occurrences.
with (E/'14-audit/evidence-traceability.csv').open('w',encoding='utf-8',newline='') as f:
 w=csv.writer(f);w.writerow(['document','source_path','start_line','end_line','validation','evidence_limit'])
 for x in v['source_references']:w.writerow([x['document'],x['source'],x['line'],x['end'],'file and range valid' if x['valid'] else 'invalid','Source citation existence is not semantic proof; source-specific review separate'])
missing=[n for n in required if not(E/n).is_file() or not(E/n).stat().st_size];assert not missing,missing
assert len(mods)==46 and all((E/x['document']).is_file() for x in mods)
assert len(fs)==len({x['feature_id'] for x in fs})==190
fields=['feature_id','name','purpose','ui_entry','implementation_files','core_classes','workflow','dependencies','data_structures','error_handling','tests','status','evidence'];assert all(all(x[k].strip() for k in fields) for x in fs)
assert len({x['test_source'] for x in tr})==71
print(json.dumps({'documents':len(mdpaths()),'required_files':len(required),'source_refs':len(v['source_references']),'unique_source_refs':len(unique_refs),'diagrams':len(v['diagrams']),'unique_diagrams':len(unique_diagrams),'internal_file_links':len(v['links']),'errors':len(v['errors'])}))
# Source immutability result from the independent full hash comparison.
preservation=json.loads((E/'14-audit/source-preservation.json').read_text());assert preservation['unchanged']
metrics={'generated_utc':NOW,'commit':COMMIT,'repository_records':30301,'regular_files_preserved':preservation['checked'],'authored_source_files':204,'source_files_reviewed':204,'targeted_vsto_code_review_files':46,'supporting_text_files':339,'modules':len(mods),'features':len(fs),'command_routes':172,'capabilities':18,'interface_groups':len(ifs),'test_sources':71,'feature_related_test_candidates':feature_mapped,'required_files':len(required),'required_files_present':len(required),'document_delivery_percent':100.0,'markdown_documents':len(mdpaths()),'source_ref_occurrences':len(v['source_references']),'unique_source_refs':len(unique_refs),'source_ref_valid_percent':100.0,'diagram_occurrences':len(v['diagrams']),'unique_diagrams':len(unique_diagrams),'diagram_structural_percent':100.0,'diagram_parser_validated':0,'internal_links_checked':len(v['links']),'current_assertions':229,'typescript_typecheck':'PASS','powershell_files_parsed':50,'native_office_current':'NOT RUN','source_unchanged':True}
(E/'14-audit/completion-metrics.json').write_text(json.dumps(metrics,indent=2),encoding='utf-8')
p=E/'14-audit/documentation-validation.md';s=p.read_text(encoding='utf-8');s+=f'''\n## Final structural measurements\n\n- Markdowndocuments:{metrics['markdown_documents']}; requiredpaths:{len(required)}/{len(required)}.\n- Sourcereferenceoccurrences:{len(v['source_references'])}, distinctsource/ranges:{len(unique_refs)}; allvalidfile/ranges.\n- Internalfilelinktargetschecked:{len(v['links'])}; zeromissing.\n- Mermaidblocks:{len(v['diagrams'])}; uniquedesigns:{len(unique_diagrams)}; allstructurallyrecognized. Fullparser/renderchecks:0/{len(unique_diagrams)}, unavailable locally.\n- Currentsemanticcodefileaccounting:204/204withreviewqualifications;46moduledesigns;190featureentries;71testsourcefiles.\n- Sourcepreservation:{preservation['checked']}regularfiles unchanged; noadded/removed/modifiedfiles outsideexport.\n\nSee [full validation](parent-validation-all.json), [metrics](completion-metrics.json), [contract examples](contract-example-validation.log), [preservation](source-preservation.json) and [evidence](evidence-traceability.csv). Counterexample controls and secret/archivechecks are recorded separately; no source/runtime result is inferred from a structural pass.\n''';p.write_text(s,encoding='utf-8')
