"""Repair a new, offline PPTX copy. Never edits the input or an open presentation."""
import argparse, hashlib, json, re
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
A='{'+NS['a']+'}'
def bounds(shape):
    off=shape.find('p:spPr/a:xfrm/a:off',NS);ext=shape.find('p:spPr/a:xfrm/a:ext',NS)
    if off is None or ext is None:return None
    return tuple(int(e.get(k)) for e,k in [(off,'x'),(off,'y'),(ext,'cx'),(ext,'cy')])
def has_text(shape):return bool(shape.findall('.//a:t',NS))
def contains(box,frame):
    x,y,w,h=box;tx,ty,tw,th=frame
    return x<=tx+tw/2<=x+w and y<=ty+th/2<=y+h

def fix_slide(data):
    root=E.fromstring(data,E.XMLParser(resolve_entities=False,no_network=True));expanded=0
    for parent in root.xpath('//p:spTree | //p:grpSp',namespaces=NS):
        shapes=parent.findall('p:sp',NS);labels=[s for s in shapes if has_text(s) and bounds(s)]
        containers=[s for s in shapes if not has_text(s) and bounds(s) and s.find('p:spPr/a:solidFill',NS) is not None]
        assigned=set()
        for box in sorted(containers,key=lambda s:bounds(s)[2]*bounds(s)[3]):
            b=bounds(box);inside=[s for s in labels if id(s) not in assigned and contains(b,bounds(s))]
            if len(inside)!=1:continue
            text=inside[0];x,y,w,h=b;geometry=box.find('p:spPr/a:prstGeom',NS);kind=geometry.get('prst') if geometry is not None else 'freeform'
            px=int(w*(.23 if kind in ('chevron','diamond','freeform') else .06));py=int(h*.12)
            off=text.find('p:spPr/a:xfrm/a:off',NS);ext=text.find('p:spPr/a:xfrm/a:ext',NS)
            off.set('x',str(x+px));off.set('y',str(y+py));ext.set('cx',str(max(1,w-2*px)));ext.set('cy',str(max(1,h-2*py)));expanded+=1;assigned.add(id(text))
    bodies=root.xpath('//p:txBody | //a:txBody',namespaces=NS)
    for body in bodies:
        prop=body.find(A+'bodyPr')
        if prop is None:prop=E.Element(A+'bodyPr');body.insert(0,prop)
        prop.set('anchor','ctr')
        for paragraph in body.findall(A+'p'):
            pp=paragraph.find(A+'pPr')
            if pp is None:pp=E.Element(A+'pPr');paragraph.insert(0,pp)
            pp.set('algn','ctr')
    return E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True),len(bodies),expanded

def repair(source,output):
    source=Path(source).resolve();output=Path(output).resolve()
    if source==output or output.exists():raise ValueError('Choose a new output path; existing files are never overwritten.')
    output.parent.mkdir(parents=True,exist_ok=True)
    original_hash=hashlib.sha256(source.read_bytes()).hexdigest();slides=fields=expanded=0
    with ZipFile(source) as zin,ZipFile(output,'w') as zout:
        for info in zin.infolist():
            data=zin.read(info.filename)
            if re.fullmatch(r'ppt/slides/slide\d+\.xml',info.filename):data,count,fit=fix_slide(data);slides+=1;fields+=count;expanded+=fit
            zout.writestr(info,data)
    # Compare every untouched ZIP entry and every text value. No package parts removed.
    with ZipFile(source) as before,ZipFile(output) as after:
        assert before.namelist()==after.namelist()
        for name in before.namelist():
            if re.fullmatch(r'ppt/slides/slide\d+\.xml',name):
                old=E.fromstring(before.read(name));new=E.fromstring(after.read(name))
                assert old.xpath('//a:t/text()',namespaces=NS)==new.xpath('//a:t/text()',namespaces=NS)
                for tag in ('timing','transition'):
                    x=old.find('p:'+tag,NS);y=new.find('p:'+tag,NS)
                    assert (E.tostring(x) if x is not None else None)==(E.tostring(y) if y is not None else None)
                assert all(p.get('anchor')=='ctr' for p in new.xpath('//p:txBody/a:bodyPr | //a:txBody/a:bodyPr',namespaces=NS))
                assert all(p.get('algn')=='ctr' for p in new.xpath('//p:txBody/a:p/a:pPr | //a:txBody/a:p/a:pPr',namespaces=NS))
            else:assert before.read(name)==after.read(name)
    assert hashlib.sha256(source.read_bytes()).hexdigest()==original_hash
    report={'slides':slides,'textFramesCentered':fields,'singleLabelsFitted':expanded,'inputUnchanged':True,'textAndOtherPartsPreserved':True,'powerPointRenderingVerified':False}
    output.with_suffix('.validation.json').write_text(json.dumps(report,indent=2))
    print(json.dumps(report))
if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--input',required=True);parser.add_argument('--output',required=True);args=parser.parse_args();repair(args.input,args.output)
