from pathlib import Path
import json,csv,re,hashlib,os,zipfile,datetime,importlib.util,sys,subprocess,collections
R=Path(__file__).resolve().parents[2];E=R/'project-export';A=E/'14-audit';D=E/'15-deliverables'
spec=importlib.util.spec_from_file_location('export_validator',E/'_analysis/validate-export.py');v=importlib.util.module_from_spec(spec);spec.loader.exec_module(v)
# Editorial spacing outside fenced contracts/diagrams and protected inline code/link targets.
fixes={'authoredcode':'authored code','sourcefiles':'source files','fulltext':'full-text','requiredfilelist':'required file list','requiredpaths':'required paths','moduledesigns':'module designs','sourcecitations':'source citations','read-onlypath':'read-only path','scannerclearance':'scanner clearance','fullMermaidparser':'full Mermaid parser','fullparser':'full parser','fullsyntax':'full syntax','document-delivery':'document delivery','fullsemantic':'full semantic','nativeOffice':'native Office','signedinstaller':'signed installer','Officeactivation':'Office activation','runtimeacceptance':'runtime acceptance','checksareperformed':'checks are performed','allbehavior':'all behavior','sourceprovenance':'source provenance','privateconfiguration':'private configuration','licensedtemplateassets':'licensed template assets','sourcepaths':'source paths','linevalidation':'line validation','fence/type':'fence/type','filelinktargetschecked':'file link targets checked','fullbaseline':'full baseline','onlyinsideexport':'only inside export','before/afterbyte':'before/after byte','uniquedesigns':'unique designs','Sourcereferenceoccurrences':'Source reference occurrences','Internalfilelinktargetschecked':'Internal file link targets checked','Mermaidblocks':'Mermaid blocks','allstructurallyrecognized':'all structurally recognized','Currentsemanticcodefileaccounting':'Current semantic code-file accounting','Sourcepreservation':'Source preservation','noadded/removed/modifiedfiles':'no added/removed/modified files','outsideexport':'outside export','Exportversion':'Export version','source/config/dependency':'source/configuration/dependency','NoAPI':'No API','signedinstallation':'signed installation','fullbody':'full-body','testsourcefiles':'test source files','testcandidates':'test candidates','notinstrumentedcoverage':'not instrumented coverage','notproduction':'not production','datafile':'data file','liveAI':'live AI','livedeployment':'live deployment','testfixtures':'test fixtures'}
for p in [x for x in E.rglob('*.md') if '_analysis' not in x.parts and '.unlazy' not in x.parts]:
 text=p.read_text(encoding='utf-8-sig');out=[];fence=False
 for line in text.splitlines():
  if line.lstrip().startswith('```'):fence=not fence;out.append(line);continue
  if not fence:
   parts=re.split(r'(`[^`]*`|\[[^\]\n]*\]\([^\n)]*\))',line)
   for i in range(0,len(parts),2):
    t=re.sub(r'(?<![\w./:@-])(\d+)(?=[A-Za-z])',r'\1 ',parts[i])
    for old,new in fixes.items():t=re.sub(r'\b'+re.escape(old)+r'\b',lambda m:new,t)
    parts[i]=t
   line=''.join(parts)
  out.append(line)
 p.write_text('\n'.join(out)+'\n',encoding='utf-8')
# Prove structural validator detects real counterexamples before using absence results.
control=E/'_analysis/validator-controls';control.mkdir(exist_ok=True)
good=control/'good.md';bad=control/'bad.md';good.write_text('# Good\n\n`apps/powerpoint-vsto/Contracts.cs:7`\n\n```mermaid\nflowchart LR\n A[One] --> B[Two]\n```\n',encoding='utf-8');bad.write_text('# Bad\n\n`apps/powerpoint-vsto/Contracts.cs:999999`\n\n[Missing](missing.md)\n\n```mermaid\ninvalidDiagram\n```\n',encoding='utf-8')
assert not v.validate([good],True)['errors'];badresult=v.validate([bad],True);assert len(badresult['errors'])==3
secret_patterns=[re.compile(r'-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----'),re.compile(r'\b(?:AKIA|ASIA)[A-Z0-9]{16}\b'),re.compile(r'\bgh[pousr]_[A-Za-z0-9]{30,}\b'),re.compile(r'\bsk-[A-Za-z0-9_-]{30,}\b'),re.compile(r'\bxox[baprs]-[A-Za-z0-9-]{25,}\b')]
def secret_hits(text):return [i for i,p in enumerate(secret_patterns) if p.search(text)]
assert secret_hits('-----BEGIN '+'PRIVATE KEY-----') and not secret_hits('Use a nonsecret approved certificate thumbprint.')
(A/'validator-controls.json').write_text(json.dumps({'positive_control':'PASS source/ref+diagram','negative_control':'PASS invalid source/missing link/bad diagram type rejected','secret_positive_control':'PASS synthetic marker detected','secret_negative_control':'PASS ordinary documentation not flagged','not_a_mermaid_parser':True},indent=2),encoding='utf-8')
# Selected portable files; validation machinery/source/binaries/assets never included.
allowed={'.md','.csv','.json','.log','.sha256'}
exclude_names={'export-file-inventory.csv','archive-validation.json','DOCUMENT-CHECKSUMS.sha256','PowerPoint-Intelligence-NexWave-Technical-Design.zip.sha256'}
def portable():
 return sorted([p for p in E.rglob('*') if p.is_file() and p.suffix.lower() in allowed and '_analysis' not in p.relative_to(E).parts and '.unlazy' not in p.relative_to(E).parts and p.name not in exclude_names],key=lambda p:p.relative_to(E).as_posix())
hits=[]
for p in portable():
 ids=secret_hits(p.read_text(encoding='utf-8-sig',errors='replace'))
 if ids:hits.append({'file':p.relative_to(E).as_posix(),'pattern_ids':ids})
if hits:raise ValueError('Potential secret markers found: '+str(hits))
(A/'secret-review.json').write_text(json.dumps({'checked_files':len(portable()),'high_confidence_hits':hits,'positive_negative_controls_passed':True,'limitations':'Pattern scan plus editorial review; not a full secret-detection product or proof all repository secrets are absent. Sensitive source/config/key files were not exported.'},indent=2),encoding='utf-8')
# First inventories/checksums enable all target links; then final validator will refresh them.
def create_inventory_and_checksums():
 ps=portable();lines=[]
 with (A/'export-file-inventory.csv').open('w',encoding='utf-8',newline='') as f:
  w=csv.writer(f);w.writerow(['export_path','bytes','sha256','scope'])
  for p in ps:
   rel=p.relative_to(E).as_posix();sha=hashlib.sha256(p.read_bytes()).hexdigest();w.writerow([rel,p.stat().st_size,sha,'portable documentation/data']);lines.append(sha+'  '+rel)
 lines.append(hashlib.sha256((A/'export-file-inventory.csv').read_bytes()).hexdigest()+'  14-audit/export-file-inventory.csv')
 (D/'DOCUMENT-CHECKSUMS.sha256').write_text('\n'.join(lines)+'\n',encoding='utf-8')
create_inventory_and_checksums()
archive=D/'PowerPoint-Intelligence-NexWave-Technical-Design.zip';archive_sha=Path(str(archive)+'.sha256')
# Native anchor check without pretending to parse Mermaid.
def slug(s):return re.sub(r'[^\w\- ]','',s.replace('`','').lower()).replace(' ','-')
def anchors(p):
 counts=collections.Counter();out=set();fence=False
 for line in p.read_text(encoding='utf-8-sig').splitlines():
  if line.lstrip().startswith('```'):fence=not fence;continue
  if fence:continue
  m=re.match(r'^#{1,6}\s+(.+)',line)
  if m:
   a=slug(m[1]);n=counts[a];out.add(a+(('-'+str(n)) if n else ''));counts[a]+=1
 return out
anchor_errors=[]
for p in v.E.rglob('*.md'):
 if '_analysis' in p.parts or '.unlazy' in p.parts:continue
 for target in re.findall(r'\[[^\]\n]*\]\(([^\n)]+)\)',p.read_text(encoding='utf-8-sig')):
  if target.startswith(('http:','https:','mailto:')) or '#' not in target:continue
  name,frag=target.split('#',1);dest=(p.parent/name).resolve() if name else p
  if dest.is_file() and dest.suffix=='.md' and frag not in anchors(dest):anchor_errors.append({'doc':p.relative_to(E).as_posix(),'target':target})
(A/'anchor-validation.json').write_text(json.dumps({'errors':anchor_errors,'algorithm':'GitHub-style heading slug with duplicate suffixes; native viewer variations remain possible'},indent=2),encoding='utf-8')
if anchor_errors:print(json.dumps(anchor_errors));raise SystemExit('Anchor validation failed')
# Write once, validate contents independently, record outside archive, then freeze final checksums.
def make_archive():
 if archive.exists():
  old=E/'_analysis'/('prior-zip-'+datetime.datetime.now().strftime('%H%M%S%f')+'.zip');archive.replace(old)
 ps=portable()+[A/'export-file-inventory.csv',D/'DOCUMENT-CHECKSUMS.sha256']
 with zipfile.ZipFile(archive,'x',zipfile.ZIP_DEFLATED,compresslevel=9) as z:
  for p in ps:z.write(p,'project-export/'+p.relative_to(E).as_posix())
 return ps
create_inventory_and_checksums();ps=make_archive()
sha=hashlib.sha256(archive.read_bytes()).hexdigest();archive_sha.write_text(sha+'  '+archive.name+'\n',encoding='utf-8')
with zipfile.ZipFile(archive) as z:
 names=z.namelist();assert len(names)==len(set(names));assert z.testzip() is None
 for p in ps:
  key='project-export/'+p.relative_to(E).as_posix();assert hashlib.sha256(z.read(key)).digest()==hashlib.sha256(p.read_bytes()).digest()
 assert not any(Path(n).suffix.lower() in {'.cs','.ps1','.py','.mjs','.ts','.exe','.dll','.pfx','.p12','.config','.pptx','.potx','.docx','.zip'} for n in names)
 assert not any('/_analysis/' in n or '/.unlazy/' in n or '..' in Path(n).parts for n in names)
result={'archive':archive.relative_to(E).as_posix(),'files':len(ps),'bytes':archive.stat().st_size,'sha256':sha,'crc_verified':True,'all_entry_hashes_match_current_files':True,'application_source_private_config_binaries_assets_excluded':True,'self_inclusion':False}
(A/'archive-validation.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
# The archive-validation/hash files intentionally stay external to avoid self-referential archive hashing.
proc=subprocess.run([sys.executable,'-B',str(E/'_analysis/validate-export.py'),'all'],capture_output=True,text=True,encoding='utf-8');(A/'export-validation-transcript.log').write_text(proc.stdout+proc.stderr,encoding='utf-8');print(proc.stdout)
if proc.returncode:raise SystemExit('Final document link/reference validation failed')
# Freeze the now-final validation JSON/transcript and metrics into the archive.
measured=json.loads((A/'parent-validation-all.json').read_text(encoding='utf-8'))
mp=A/'completion-metrics.json';mm=json.loads(mp.read_text(encoding='utf-8'));mm['internal_links_checked']=len(measured['links']);mm['validation_errors']=len(measured['errors']);mm['portable_documentation_files']=len(ps);mp.write_text(json.dumps(mm,indent=2),encoding='utf-8')
dp=A/'documentation-validation.md';dt=dp.read_text(encoding='utf-8');dt=re.sub(r'Internal file link targets checked:[0-9]+', 'Internal file link targets checked:'+str(len(measured['links'])), dt);dt=re.sub(r'[0-9]+ internal file links',str(len(measured['links']))+' internal file links',dt);dp.write_text(dt,encoding='utf-8')
create_inventory_and_checksums();ps=make_archive()
sha=hashlib.sha256(archive.read_bytes()).hexdigest();archive_sha.write_text(sha+'  '+archive.name+'\n',encoding='utf-8')
with zipfile.ZipFile(archive) as z:
 assert z.testzip() is None
 for p in ps:assert hashlib.sha256(z.read('project-export/'+p.relative_to(E).as_posix())).digest()==hashlib.sha256(p.read_bytes()).digest()
result.update(files=len(ps),bytes=archive.stat().st_size,sha256=sha)
(A/'archive-validation.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps({'archive':result,'structural_and_negative_controls':'PASS','anchor_errors':len(anchor_errors),'secret_hits':len(hits)}))
