"""Repair a new, offline PPTX copy. Never edits the input or an open presentation."""
import argparse
from lxml import etree as E
import RepairTemplateSafety as SAFETY
from RepairTemplateSafety import Limits, Budget, require
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

def fix_slide(data,budget=None):
    budget=budget or Budget(Limits())
    root=SAFETY.parse_xml(data,budget);expanded=0
    for parent in root.xpath('//p:spTree | //p:grpSp',namespaces=NS):
        shapes=parent.findall('p:sp',NS);labels=[s for s in shapes if has_text(s) and bounds(s)]
        containers=[s for s in shapes if not has_text(s) and bounds(s) and s.find('p:spPr/a:solidFill',NS) is not None]
        assigned=set()
        for box in sorted(containers,key=lambda s:bounds(s)[2]*bounds(s)[3]):
            budget.check(work=len(labels)+1)
            b=bounds(box);inside=[s for s in labels if id(s) not in assigned and contains(b,bounds(s))]
            if len(inside)!=1:continue
            text=inside[0];x,y,w,h=b;geometry=box.find('p:spPr/a:prstGeom',NS);kind=geometry.get('prst') if geometry is not None else 'freeform'
            px=int(w*(.23 if kind in ('chevron','diamond','freeform') else .06));py=int(h*.12)
            off=text.find('p:spPr/a:xfrm/a:off',NS);ext=text.find('p:spPr/a:xfrm/a:ext',NS)
            off.set('x',str(x+px));off.set('y',str(y+py));ext.set('cx',str(max(1,w-2*px)));ext.set('cy',str(max(1,h-2*py)));expanded+=1;assigned.add(id(text))
    bodies=root.xpath('//p:txBody | //a:txBody',namespaces=NS)
    for body in bodies:
        budget.check(work=1)
        prop=body.find(A+'bodyPr')
        if prop is None:prop=E.Element(A+'bodyPr');body.insert(0,prop)
        prop.set('anchor','ctr')
        for paragraph in body.findall(A+'p'):
            budget.check(work=1)
            pp=paragraph.find(A+'pPr')
            if pp is None:pp=E.Element(A+'pPr');paragraph.insert(0,pp)
            pp.set('algn','ctr')
    return E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True),len(bodies),expanded

def repair(source,output,limits=None):
    return SAFETY.repair(source,output,fix_slide,limits)
if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--input',required=True);parser.add_argument('--output',required=True);args=parser.parse_args();repair(args.input,args.output)
