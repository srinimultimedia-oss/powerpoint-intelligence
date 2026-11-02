"""Read-only repository content audit; outputs inventory and classified dependency hits."""
import argparse, csv, hashlib, json, os, re, subprocess, zipfile
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--root',required=True);p.add_argument('--git',required=True);a=p.parse_args()
root=Path(a.root).resolve()
def git(*args):
 return subprocess.check_output([a.git,'-C',str(root),*args]).decode('utf-8','replace')
tracked=set(git('ls-files','-z').rstrip('\0').split('\0'))
baseline=git('rev-parse','HEAD').strip()
paths=sorted((p for p in root.rglob('*') if p.is_file()),key=lambda p:str(p).lower())
# Every file is recorded, including ignored/generated/admin data. No binary is claimed as source-reviewed.
relpaths=[str(x.relative_to(root)).replace('\\','/') for x in paths]
ignored=set(subprocess.run([a.git,'-C',str(root),'check-ignore','--stdin','-z'],input=('\0'.join(relpaths)+'\0').encode(),stdout=subprocess.PIPE,check=False).stdout.decode('utf-8','replace').strip('\0').split('\0'))
pattern=re.compile(r'\.tcidx|tcidx|think[- ]?cell|library index|slide index|slide library|template library|presentation library',re.I)
loading=re.compile(r'DllImport|Assembly\.Load|Process\.Start|GetTypeFromProgID|GetTypeFromCLSID|Registry\.|HintPath|COMReference|TemplateLibraryFolder|LocalLibraryCache|LocalSlideIndex')
text_ext={'.cs','.csproj','.ps1','.json','.md','.txt','.yml','.yaml','.config','.xml','.xsd','.mjs','.ts','.csv','.py','.gitignore'}
rows=[];hits=[];errors=[]
for path,rel in zip(paths,relpaths):
 admin=rel.startswith('.git/');generated=admin or any(x in {'artifacts','bin','obj','node_modules','.vs'} for x in path.relative_to(root).parts)
 scope='Tracked' if rel in tracked else 'Git administrative' if admin else 'Ignored' if rel in ignored else 'Untracked'
 module=rel.split('/')[0];inspection='';matches=[];category='G';required='No';size=0;digest='';issue=''
 try:
  raw=path.read_bytes();size=len(raw);digest=hashlib.sha256(raw).hexdigest().upper()
  if not admin and rel not in {'docs/migration/repository-files.csv','docs/migration/dependency-findings.csv'} and (path.suffix.lower() in text_ext or path.name in {'.gitignore','AGENTS.md'}):
   try:content=raw.decode('utf-8-sig');encoding='UTF-8'
   except UnicodeDecodeError:content=raw.decode('cp1252');encoding='Windows-1252 fallback explicitly recorded'
   inspection='Entire text decoded ('+encoding+') and dependency/loading patterns scanned'
   for n,line in enumerate(content.splitlines(),1):
    found=pattern.search(line);load=loading.search(line)
    if found or load:
     token=(found or load).group(0)
     c='E' if module=='docs' or path.name in {'README.md','AGENTS.md'} else 'F' if module=='tests' or path.name.startswith('Test-') else 'C' if path.suffix=='.config' else 'B' if path.suffix in {'.csproj','.ps1','.yml','.yaml'} else 'A' if module in {'apps','src'} else 'G'
     if path.name in {'DemoRibbon.cs','LocalLibraryGallery.cs'}:c='D'
     if path.name=='FeatureInventory.cs' and found:c='G' # descriptive demo inventory, no service invocation
     hits.append(dict(File=rel,Line=n,Classification=c,Token=token));matches.append(c)
   if matches:category=','.join(sorted(set(matches)))
   if path.name in {'TemplateLibrary.cs','LocalLibraryGallery.cs','PowerPointIntelligence.csproj','App.config'}:required='Yes'
   elif matches and module in {'scripts','tests'}:required='Review integration scope'
  elif not admin and path.suffix.lower() in {'.pptx','.potx','.docx'}:
   with zipfile.ZipFile(path) as z:
    names=z.namelist();inspection='ZIP member inventory and package signatures only; '+str(len(names))+' members'
    if any('tcidx' in x.lower() for x in names):category='G';issue='Index-named member requires classification'
  else:inspection='Byte hash, file size/type metadata only; not manually reviewed or executed'
 except Exception as ex:
  issue=type(ex).__name__+': '+str(ex);inspection='Inspection failed';errors.append(dict(File=rel,Error=issue))
 rows.append(dict(File=rel,Scope=scope,Type=path.suffix or 'extensionless',Module=module,Bytes=size,SHA256=digest,Inspection=inspection,Classification=category,ModificationRequired=required,Modified='Not yet assessed',Verification='Inventory scan only; functional verification pending',Outstanding=issue))
out=root/'docs'/'migration';out.mkdir(parents=True,exist_ok=True)
for name,data,keys in [('repository-files.csv',rows,list(rows[0])),('dependency-findings.csv',hits,['File','Line','Classification','Token'])]:
 with (out/name).open('w',newline='',encoding='utf-8') as f:
  w=csv.DictWriter(f,fieldnames=keys);w.writeheader();w.writerows(data)
missing=sorted(tracked-set(relpaths))
summary=dict(Baseline=baseline,Tracked=len(tracked),InventoriedTracked=sum(x['Scope']=='Tracked' for x in rows),TotalFiles=len(rows),Untracked=sum(x['Scope']=='Untracked' for x in rows),Ignored=sum(x['Scope']=='Ignored' for x in rows),Administrative=sum(x['Scope']=='Git administrative' for x in rows),Findings=len(hits),MissingTracked=missing,InspectionErrors=errors)
(out/'inventory-summary.json').write_text(json.dumps(summary,indent=2),encoding='utf-8')
md=['# Repository inventory baseline','',f'Baseline commit: `{baseline}`. Audit includes the merged untracked source; Git alone contains only {len(tracked)} files.','',f'Accounted files: {len(rows)}. Tracked present: {summary["InventoriedTracked"]}/{len(tracked)}. Inspection errors: {len(errors)}.','', 'The companion [repository-files.csv](repository-files.csv) records every enumerated file, scope, module, SHA256, inspection method, classification and pending verification. [dependency-findings.csv](dependency-findings.csv) records classified line locations. Binary package member inventories are not proof of native behavior; generated output and Git administrative data are metadata-only. No Office test was executed by this audit.','', 'This is the pre-implementation baseline. Newly generated audit files and subsequent changes require inclusion in the final coverage pass. Existing modified/untracked implementation is owner baseline, not attributed to this migration.']
(out/'REPOSITORY_FILE_INVENTORY.md').write_text('\n'.join(md)+'\n',encoding='utf-8')
print(json.dumps(summary,indent=2))
if missing or errors:raise SystemExit(1)