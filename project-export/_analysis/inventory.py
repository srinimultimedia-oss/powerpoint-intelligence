from pathlib import Path
import os,json,csv,hashlib,re,datetime,subprocess,zipfile,xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];OUT=ROOT/'project-export'
TEXT={'.cs','.ps1','.py','.mjs','.ts','.tsx','.csproj','.config','.json','.xsd','.xml','.yml','.yaml','.md','.txt','.csv'}
CODE={'.cs','.ps1','.py','.mjs','.ts','.tsx'}
secret={'.pfx','.p12','.snk','.key','.pem'}
rows=[];baseline={};symbols=[];texts={};errors=[]
def hashfile(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for chunk in iter(lambda:f.read(1048576),b''):h.update(chunk)
 return h.hexdigest().upper()
def classify(rel):
 a=rel.parts;low=[x.lower() for x in a];ext=rel.suffix.lower()
 if a[0]=='.git':return 'Git metadata','excluded-vcs-metadata'
 if a[0]=='artifacts' or any(x in low for x in ['bin','obj','node_modules','__pycache__']):return 'Generated/vendor evidence','excluded-generated-or-vendor'
 if ext in secret or rel.name.startswith('.env'):return 'Sensitive material','excluded-sensitive-no-content-read'
 if ext in CODE:return '/'.join(a[:2]) if len(a)>1 else 'Root','authored-code'
 if ext in TEXT:return '/'.join(a[:2]) if len(a)>1 else 'Root','configuration-schema-documentation'
 if ext in {'.potx','.pptx','.docx'}:return 'Document/template asset','binary-package-inventory-only'
 return 'Asset/dependency/binary','binary-inventory-only'
for base,dirs,files in os.walk(ROOT,followlinks=False):
 dirs[:]=sorted(d for d in dirs if d!='project-export')
 for d in list(dirs):
  p=Path(base)/d
  if p.is_symlink() or (getattr(p.lstat(),'st_file_attributes',0)&1024):
   dirs.remove(d);rows.append({'path':p.relative_to(ROOT).as_posix(),'extension':'<directory-link>','bytes':0,'module':'Reparse point','analysis_status':'excluded-reparse-not-followed','sha256':'','lines':0})
 for name in sorted(files):
  p=Path(base)/name;rel=p.relative_to(ROOT);mod,status=classify(rel)
  try:
   st=p.lstat()
   if p.is_symlink() or (getattr(st,'st_file_attributes',0)&1024):status='excluded-reparse-not-followed';sha=''
   else:sha=hashfile(p);baseline[rel.as_posix()]={'sha256':sha,'bytes':st.st_size}
   row={'path':rel.as_posix(),'extension':p.suffix.lower() or '<none>','bytes':st.st_size,'module':mod,'analysis_status':status,'sha256':sha,'lines':0}
   if status in ['authored-code','configuration-schema-documentation']:
    raw=p.read_bytes()
    try:text=raw.decode('utf-8-sig')
    except UnicodeDecodeError:text=raw.decode('cp1252',errors='replace')
    texts[rel.as_posix()]=text;row['lines']=len(text.splitlines())
    row['analysis_status']='structurally-inspected-awaiting-semantic-review'
    for i,line in enumerate(text.splitlines(),1):
     if (re.search(r'\b(class|interface|enum|struct|record)\s+[A-Za-z_]|\bfunction\s+[\w-]+|\bexport\s+(?:async\s+)?(?:function|class|interface|type|const)\s+|\b(?:public|internal|private|protected)\s+(?:(?:static|sealed|virtual|override|async|readonly|partial)\s+)*[\w<>\[\],.?]+\s+\w+\s*\(',line)):
      symbols.append({'path':rel.as_posix(),'line':i,'declaration':line.strip()[:1500],'extraction':'lexical-candidate-not-AST'})
   rows.append(row)
  except Exception as exc:errors.append({'path':rel.as_posix(),'error':type(exc).__name__});rows.append({'path':rel.as_posix(),'extension':p.suffix,'bytes':0,'module':mod,'analysis_status':'unreadable','sha256':'','lines':0})
  if len(rows)%10000==0:print('inventoried',len(rows),flush=True)
with (OUT/'01-repository-analysis/complete-file-inventory.csv').open('w',encoding='utf-8',newline='') as f:
 w=csv.DictWriter(f,fieldnames=['path','extension','bytes','module','analysis_status','sha256','lines']);w.writeheader();w.writerows(rows)
with (OUT/'01-repository-analysis/symbol-inventory.csv').open('w',encoding='utf-8',newline='') as f:
 w=csv.DictWriter(f,fieldnames=['path','line','declaration','extraction']);w.writeheader();w.writerows(symbols)
(OUT/'_analysis/baseline.json').write_text(json.dumps(baseline,separators=(',',':')),encoding='utf-8')
(OUT/'_analysis/text-index.json').write_text(json.dumps({k:{'lines':len(v.splitlines()),'bytes':len(v.encode('utf-8'))} for k,v in texts.items()},indent=2),encoding='utf-8')
(OUT/'_analysis/discovery.json').write_text(json.dumps({'time':datetime.datetime.now(datetime.timezone.utc).isoformat(),'files':len(rows),'bytes':sum(x['bytes'] for x in rows),'authored_code':sum(x['extension'] in CODE and x['path'] in texts for x in rows),'text_files':len(texts),'symbols':len(symbols),'errors':errors},indent=2),encoding='utf-8')
print(json.dumps({'files':len(rows),'text_files':len(texts),'symbols':len(symbols),'errors':errors}))
