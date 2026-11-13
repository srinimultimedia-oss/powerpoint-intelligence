from pathlib import Path
import re,json,csv,hashlib
r=Path(__file__).resolve().parents[2]; e=r/'project-export'; out=e/'08-data-design'; mods=e/'03-modules/individual-module-designs/core'
def read(n): return (r/n).read_text(encoding='utf-8-sig')
def ref(n,token=''):
 return '`'+n+':'+str(next((i for i,s in enumerate(read(n).splitlines(),1) if token in s),1))+'`'
def write(n,s):
 p=e/n;p.parent.mkdir(parents=True,exist_ok=True);p.write_text(s.strip()+'\n',encoding='utf-8')
def table(headers,rows):
 return '| '+' | '.join(headers)+' |\n| '+' | '.join(['---']*len(headers))+' |\n'+'\n'.join('| '+' | '.join(str(v).replace('|','\\|').replace('\n',' ') for v in row)+' |' for row in rows)+'\n'
def module(mid,title,files,purpose,status,io,deps,flow,errors,config,security,tests,limits,debt,extension):
 symbols=[];ops=[]
 for f in files:
  for i,line in enumerate(read(f).splitlines(),1):
   if re.search(r'\b(class|interface|enum|type)\s+[A-Za-z_][A-Za-z_0-9]*',line):symbols.append(ref(f,line.strip())+' — '+line.strip())
   if re.search(r'(?:public|internal|export)\s+(?:static\s+|async\s+)?[^;{}=]*\([^;{}]*\)',line) and not line.strip().startswith('//'):ops.append(ref(f,line.strip())+' — '+line.strip())
 body='# '+title+'\n\n**VERIFIED source review.** This design describes inspected code, not current runtime/Office execution.\n\n'
 sections=[('1. Purpose',purpose),('2. Implementation status',status),('3. Source locations','\n'.join('- '+ref(f) for f in files)),('4. Internal components',deps+' '+flow),('5. Classes and interfaces','\n'.join('- '+s for s in symbols) or 'Static exported functions only; no class type in these files.'),('6. Important methods and functions','Declaration locator below is lexical and may include overload/call-like lines; semantic workflow is described separately.\n\n'+'\n'.join('- '+o for o in ops)),('7. Inputs and outputs',io),('8. Dependencies',deps),('9. Workflow',flow),('10. Integration points',io+' '+status),('11. Exceptions and recovery',errors),('12. Configuration dependencies',config),('13. Security considerations',security),('14. Unit and integration test definitions','\n'.join('- '+ref(t) for t in tests)+'\n\nInspected definitions only. See the parent current test report for execution; historical passes do not establish current acceptance.'),('15. Known limitations',limits),('16. Technical debt',debt),('17. Extension points',extension)]
 body+='\n\n'.join('## '+h+'\n\n'+t for h,t in sections)
 write('03-modules/individual-module-designs/core/'+mid+'.md',body)
 index=out/'core-module-map.csv';fresh=not index.exists()
 with index.open('a',encoding='utf-8',newline='') as f:
  w=csv.writer(f)
  if fresh:w.writerow(['module_id','module_name','implementation_files','test_files'])
  w.writerow([mid,title,'; '.join(files),'; '.join(tests)])
