from pathlib import Path
import subprocess,os,json,datetime,shutil,re
ROOT=Path(__file__).resolve().parents[2];E=ROOT/'project-export';OUT=E/'12-testing/current-validation';OUT.mkdir(exist_ok=True)
env=os.environ.copy();env['PYTHONDONTWRITEBYTECODE']='1';env['NODE_DISABLE_COMPILE_CACHE']='1';env.pop('NODE_COMPILE_CACHE',None)
tmp=E/'_analysis/runtime-temp';tmp.mkdir(exist_ok=True);env['TEMP']=env['TMP']=str(tmp)
results=[]
def run(name,args,cwd=ROOT):
 start=datetime.datetime.now(datetime.timezone.utc).isoformat()
 try:
  p=subprocess.run(args,cwd=cwd,env=env,capture_output=True,timeout=180)
  text=(p.stdout+p.stderr).decode('utf-8',errors='replace');code=p.returncode
 except Exception as ex:text=type(ex).__name__+': '+str(ex);code=-1
 (OUT/(name+'.log')).write_text(text,encoding='utf-8')
 counts={k:int(v) for k,v in re.findall(r'# (tests|pass|fail|skipped) (\d+)',text)}
 results.append({'name':name,'start_utc':start,'end_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'command':args,'cwd':str(cwd),'exit_code':code,'status':'PASS' if code==0 else 'FAIL','counts':counts,'log':'12-testing/current-validation/'+name+'.log'})
 print(name,code,counts,flush=True)
node=shutil.which('node.exe') or shutil.which('node')
for package in ['chart-core','chart-domain','chart-engine']:
 folder=ROOT/'packages'/package;tests=sorted(str(p.relative_to(folder)) for p in (folder/'test').glob('*.test.mjs'));run(package,[node,'--test',*tests],folder)
run('typescript-noemit',[node,str(ROOT/'packages/chart-engine/node_modules/typescript/bin/tsc'),'--noEmit','-p',str(ROOT/'packages/chart-engine/tsconfig.json')])
(OUT/'results.json').write_text(json.dumps(results,indent=2),encoding='utf-8')
