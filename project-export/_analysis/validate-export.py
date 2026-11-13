from pathlib import Path
import re,json,sys,hashlib,csv,collections
R=Path(__file__).resolve().parents[2];E=R/'project-export'
PAT=re.compile(r'(?P<path>(?:apps|src|scripts|packages|services|tests|easySetup|docs|\.github)/[A-Za-z0-9_.@/()\-]+\.(?:cs|ps1|mjs|ts|tsx|csproj|config|md|json|py|xml|xsd|yml|yaml|csv|txt)):(?P<line>\d+)(?:[-–](?P<end>\d+))?')
def validate(files,links=False):
 errors=[];refs=[];diagrams=[];linkrefs=[]
 for p in files:
  txt=p.read_text(encoding='utf-8-sig');rel=p.relative_to(E).as_posix()
  fences=re.findall(r'^\s*(`{3,}|~{3,})([^\n]*)$',txt,re.M)
  if len(fences)%2:errors.append({'doc':rel,'error':'unbalanced fenced blocks'})
  for m in PAT.finditer(txt):
   f=R/m['path'];line=int(m['line']);end=int(m['end'] or line);ok=f.is_file() and line>0 and end>=line and end<=len(f.read_text(encoding='utf-8-sig',errors='replace').splitlines());refs.append({'document':rel,'source':m['path'],'line':line,'end':end,'valid':ok})
   if not ok:errors.append({'doc':rel,'error':'invalid source reference','ref':m.group()})
  for i,m in enumerate(re.finditer(r'```mermaid\s*\n(.*?)\n```',txt,re.S),1):
   code=m.group(1);first=code.strip().splitlines()[0] if code.strip() else '';ok=bool(re.match(r'^(flowchart|graph|sequenceDiagram|classDiagram|erDiagram|stateDiagram|C4Context|C4Container|C4Component|C4Deployment|gantt|mindmap)\b',first));diagrams.append({'document':rel,'number':i,'type':first,'sha256':hashlib.sha256(code.encode()).hexdigest(),'structural_valid':ok,'parser_validated':False})
   if not ok:errors.append({'doc':rel,'error':'unknown/empty Mermaid diagram','number':i})
  if links:
   for target in re.findall(r'\[[^\]\n]*\]\(([^\n)]+)\)',txt):
    target=target.strip().strip('<>');target=target.split(' "')[0]
    if re.match(r'^(https?://|mailto:)',target):continue
    name=target.split('#')[0]
    if not name:continue
    f=(p.parent/name).resolve();ok=f.exists();linkrefs.append({'document':rel,'target':target,'valid':ok})
    if not ok:errors.append({'doc':rel,'error':'missing link target','target':target})
 return {'documents':len(files),'source_references':refs,'diagrams':diagrams,'links':linkrefs,'errors':errors,'limitations':['Mermaid parser/runtime unavailable locally; structural checks are not syntax/render validation.','Source reference checks prove file/line range, not semantic truth of prose.']}
if __name__=='__main__':
 arg=sys.argv[1] if len(sys.argv)>1 else 'all'
 if arg=='core':files=list((E/'08-data-design').glob('*.md'))+list((E/'03-modules/individual-module-designs/core').glob('*.md'))
 elif arg=='operations':files=[p for d in ['06-ai-architecture','07-api-design','09-security','11-build-deployment','03-modules/individual-module-designs/operations'] for p in (E/d).rglob('*.md')]
 elif arg=='vsto':files=[p for d in ['04-powerpoint-integration','05-features','03-modules/individual-module-designs/vsto'] for p in (E/d).rglob('*.md')]
 else:files=[p for p in E.rglob('*.md') if '.unlazy' not in p.parts and '_analysis' not in p.parts]
 result=validate(files,links=arg=='all');out=E/'14-audit'/('parent-validation-'+arg+'.json');out.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'scope':arg,'documents':result['documents'],'refs':len(result['source_references']),'diagrams':len(result['diagrams']),'errors':result['errors']}));sys.exit(bool(result['errors']))
