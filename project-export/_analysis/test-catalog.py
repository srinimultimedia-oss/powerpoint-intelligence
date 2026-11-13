from pathlib import Path
import re,json,csv,zipfile,xml.etree.ElementTree as ET
r=Path.cwd();e=r/'project-export';rows=[];review=[]
for p in sorted((r/'tests').rglob('*')):
 if p.suffix not in ['.cs','.mjs','.py'] or any(x in p.parts for x in ['bin','obj','__pycache__']):continue
 text=p.read_text(encoding='utf-8-sig');rel=p.relative_to(r).as_posix();cases=[]
 for i,line in enumerate(text.splitlines(),1):
  for m in re.finditer(r'(?:\bTest|\btest|\bCheck|\bRequire|\bPass)\s*\(\s*"([^"\r\n]+)"|\btest\s*\(\s*\x27([^\x27\r\n]+)\x27',line):
   name=next((x for x in m.groups() if x),'');cases.append({'line':i,'assertion_description':name})
  if re.search(r'^\s*def test_',line):cases.append({'line':i,'assertion_description':line.strip().split('(')[0]})
 imports=re.findall(r'^using\s+([\w.]+);',text,re.M);symbols=re.findall(r'\b(?:class|interface|enum)\s+(\w+)',text)
 writes=sorted(set(re.findall(r'\b(?:File\.(?:Write\w+|Delete|Move|Copy)|Directory\.(?:CreateDirectory|Delete)|SaveAs|\.Save\(|writeFileSync|write_text|write_bytes|mkdir|TemporaryDirectory|subprocess\.run)',text)))
 rows.append({'path':rel,'lines':len(text.splitlines()),'classes':symbols,'imports':imports,'assertions':cases,'write_indicators':writes,'inspection':'full-text lexical test-structure review; descriptions are not executed case count'})
(e/'12-testing/test-structure.json').write_text(json.dumps(rows,indent=2),encoding='utf-8')
# Requirements are reference material, not execution instructions. Extract text in memory and store only numbered requirement headings.
p=r/'docs/requirements/architecture-functional-design.docx'
with zipfile.ZipFile(p) as z:
 root=ET.fromstring(z.read('word/document.xml'));ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
 paras=[''.join(n.itertext()) for n in []]
 paras=[''.join(n.text or '' for n in para.findall('.//w:t',ns)) for para in root.findall('.//w:p',ns)]
 headings=[s for s in paras if re.search(r'\b(?:CHT|AI|DAT|ENT|EDT|INT|SLD)-\d{3}\b',s)]
(e/'01-repository-analysis/requirements-headings.json').write_text(json.dumps({'source':p.relative_to(r).as_posix(),'paragraphs':len(paras),'requirement_mentions':headings},indent=2),encoding='utf-8')
print(json.dumps({'test_source_files':len(rows),'lexical_assertion_descriptions':sum(len(x['assertions']) for x in rows),'docx_paragraphs':len(paras),'requirement_mentions':len(headings)}))
