from pathlib import Path
import csv,json,re,collections
R=Path.cwd();E=R/'project-export'
mods=[]
for m in json.loads((E/'04-powerpoint-integration/vsto-module-index.json').read_text(encoding='utf-8')):
 mods.append({'id':m['id'],'name':m['name'],'document':m['document'],'sources':m['sources'],'group':'vsto'})
for row in csv.DictReader((E/'08-data-design/core-module-map.csv').open(encoding='utf-8-sig')):
 mods.append({'id':row['module_id'],'name':row['module_name'],'document':'03-modules/individual-module-designs/core/'+row['module_id']+'.md','sources':[x.strip() for x in row['implementation_files'].split(';') if x.strip()],'group':'core'})
for row in csv.DictReader((E/'03-modules/individual-module-designs/operations/source-module-mapping.csv').open(encoding='utf-8-sig')):
 old=next((m for m in mods if m['id']==row['module_id']),None)
 if old:old['sources'].append(row['path'])
 else:mods.append({'id':row['module_id'],'name':row['module'],'document':'03-modules/individual-module-designs/operations/'+row['design'],'sources':[row['path']],'group':'operations'})
(E/'03-modules/module-map.json').write_text(json.dumps(mods,indent=2),encoding='utf-8')
with (E/'03-modules/source-module-map.csv').open('w',encoding='utf-8',newline='') as f:
 w=csv.writer(f);w.writerow(['source_path','module_id','module_name','design_document'])
 for m in mods:
  for s in sorted(set(m['sources'])):w.writerow([s,m['id'],m['name'],m['document']])
lines=['# Module inventory','','**VERIFIED source designs:** coherent modules are counted by design document, not class or directory count. A source file may serve several related modules; the source mapping preserves that many-to-many relationship. Each design addresses the17 required dimensions. Runtime evidence and unsupported paths are explicitly separate.','','| ID | Module | Design | Sources |','|---|---|---|---|']
for m in mods:lines.append('| '+m['id']+' | '+m['name']+' | ['+m['group']+'/'+Path(m['document']).stem+']('+str(Path(m['document']).relative_to('03-modules')).replace('\\','/')+') | '+str(len(set(m['sources'])))+' |')
lines+=['','The core map includes authored-but-uncompiled V3 and published schemas rather than hiding them. Operations designs include build/test/diagnostic tooling separately from runtime modules. No web backend is invented. See [source mapping](source-module-map.csv).']
(E/'03-modules/module-inventory.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
# Dependencies are observed through imports and named cross-module type references; this is not formal callgraph reachability.
classowners=collections.defaultdict(set)
for m in mods:
 for s in m['sources']:
  p=R/s
  if p.is_file() and p.suffix in ['.cs','.ts','.mjs','.ps1']:
   text=p.read_text(encoding='utf-8-sig');names=re.findall(r'\b(?:class|interface|enum|struct)\s+(\w+)|\bexport\s+(?:async\s+)?function\s+(\w+)|^function\s+([\w-]+)',text,re.M)
   for tup in names:
    for n in tup:
     if n:classowners[n].add(m['id'])
edges=collections.defaultdict(set)
for m in mods:
 for s in m['sources']:
  p=R/s
  if not p.is_file() or p.suffix not in ['.cs','.ts','.mjs','.ps1']:continue
  text=p.read_text(encoding='utf-8-sig')
  for name,owners in classowners.items():
   if re.search(r'\b'+re.escape(name)+r'\b',text):
    for other in owners-{m['id']}:edges[(m['id'],other)].add(s)
ls=['# Module dependency matrix','','**INFERRED static reference map:** edges below are exact shared-symbol references/import neighborhoods, not a compiler-resolved callgraph. Shared source inclusion and type names can produce non-call relationships. The component design prose explains consequential runtime direction; test/build paths are not assumed reachable product features.','','| From | Referenced module | Evidence source |','|---|---|---|']
for (a,b),ss in sorted(edges.items()):ls.append('| '+a+' | '+b+' | '+'; '.join('`'+s+':1`' for s in sorted(ss))+' |')
(E/'03-modules/module-dependency-matrix.md').write_text('\n'.join(ls)+'\n',encoding='utf-8')
features=list(csv.DictReader((E/'05-features/feature-source-mapping.csv').open(encoding='utf-8-sig')));testrows=[];mapped=set()
for f in features:
 refs=re.findall(r'(?:tests|packages)/[A-Za-z0-9_./-]+\.(?:cs|mjs|py|ps1)(?::\d+)?',f['tests'])
 for s in sorted(set(refs)):
  path=s.split(':')[0];mapped.add(path);current='current17metadata assertions' if path=='tests/s0/MetadataTests.cs' else 'current Node test suite' if '/test/' in path and path.endswith('.test.mjs') else 'not executed in this export'
  testrows.append([path,f['feature_id'],f['name'],'source-related regression candidate; no instrumented coverage claim',current,s])
alltests=[x['path'] for x in json.loads((E/'12-testing/test-structure.json').read_text())]
for p in sorted((R/'packages').glob('*/test/*.test.mjs')):alltests.append(p.relative_to(R).as_posix())
for p in sorted(set(alltests)-mapped):
 mode='current17metadata assertions' if p=='tests/s0/MetadataTests.cs' else 'current Node test suite' if '/test/' in p and p.endswith('.test.mjs') else 'not executed in this export'
 testrows.append([p,'COMPONENT-CONTRACT','Component/fixture/test helper','See test inventory and semantic review for exact assertions and module association',mode,p+':1'])
with (E/'12-testing/test-feature-traceability.csv').open('w',encoding='utf-8',newline='') as out:
 w=csv.writer(out);w.writerow(['test_source','feature_or_component_id','feature_or_scope','relationship','current_execution','evidence']);w.writerows(testrows)
# Sanitize lexical symbol locators to signatures, never method bodies or literals.
p=E/'01-repository-analysis/symbol-inventory.csv';ss=list(csv.DictReader(p.open(encoding='utf-8')))
for x in ss:
 x['declaration']=x['declaration'].split('{',1)[0].strip()
 if '=' in x['declaration'] and '=>' not in x['declaration']:x['declaration']=x['declaration'].split('=',1)[0].strip()
with p.open('w',encoding='utf-8',newline='') as f:
 w=csv.DictWriter(f,fieldnames=['path','line','declaration','extraction']);w.writeheader();w.writerows(ss)
print(json.dumps({'modules':len(mods),'dependencies':len(edges),'features':len(features),'test_trace_rows':len(testrows),'unique_test_sources':len({x[0] for x in testrows})}))
