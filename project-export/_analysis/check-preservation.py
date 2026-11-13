from pathlib import Path
import os,json,hashlib,datetime
R=Path(__file__).resolve().parents[2];E=R/'project-export';baseline=json.loads((E/'_analysis/baseline.json').read_text());seen=set();changed=[];added=[];errors=[]
for base,dirs,files in os.walk(R,followlinks=False):
 dirs[:]=[d for d in dirs if d!='project-export' and not (Path(base)/d).is_symlink() and not (getattr((Path(base)/d).lstat(),'st_file_attributes',0)&1024)]
 for n in files:
  p=Path(base)/n;name=p.relative_to(R).as_posix()
  if p.is_symlink() or (getattr(p.lstat(),'st_file_attributes',0)&1024):continue
  seen.add(name)
  try:
   h=hashlib.sha256()
   with p.open('rb') as f:
    for b in iter(lambda:f.read(1048576),b''):h.update(b)
   sha=h.hexdigest().upper()
   if name not in baseline:added.append(name)
   elif sha!=baseline[name]['sha256']:changed.append(name)
  except Exception as x:errors.append({'path':name,'error':type(x).__name__})
result={'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'baseline_files':len(baseline),'checked':len(seen),'modified':changed,'added':added,'missing':sorted(set(baseline)-seen),'errors':errors,'unchanged':not(changed or added or set(baseline)-seen or errors)}
(E/'14-audit/source-preservation.json').write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result));raise SystemExit(not result['unchanged'])
