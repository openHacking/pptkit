#!/usr/bin/env python3
"""Produce manual bodyPr omission and layout-only edits for autofit research."""
import sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
root=Path(sys.argv[1]);P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
with zipfile.ZipFile(root/'baseline.pptx') as z:base={i.filename:z.read(i.filename) for i in z.infolist()}
for mode in ('inherit','edit'):
 data=base.copy()
 for i in (1,3,5):
  part=f'ppt/slides/slide{i}.xml';x=E.fromstring(data[part]);body=x.find('.//p:sp/p:txBody/a:bodyPr',ns);children=[c for c in body if c.tag in {f'{{{A}}}noAutofit',f'{{{A}}}spAutoFit',f'{{{A}}}normAutofit'}];assert len(children)==1;body.remove(children[0]);data[part]=E.tostring(x,encoding='utf-8',xml_declaration=True)
 if mode=='edit':
  changes={'none':('normAutofit',{'fontScale':'60000'}),'shrink':('noAutofit',{}),'resize':('noAutofit',{})};seen=[]
  for part,blob in list(data.items()):
   if not part.startswith('ppt/slideLayouts/slideLayout') or not part.endswith('.xml'):continue
   x=E.fromstring(blob);key=x.find('p:cSld',ns).get('name').removeprefix('Target ')
   if key not in changes:continue
   b=x.find('.//p:sp/p:txBody/a:bodyPr',ns);old=[c for c in b if c.tag in {f'{{{A}}}noAutofit',f'{{{A}}}spAutoFit',f'{{{A}}}normAutofit'}];assert len(old)==1;b.remove(old[0]);tag,attrs=changes[key];E.SubElement(b,f'{{{A}}}{tag}',attrs);data[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);seen.append(key)
  assert set(seen)==set(changes)
 with zipfile.ZipFile(root/(mode+'.pptx'),'w',zipfile.ZIP_DEFLATED) as z:
  for name,value in data.items():z.writestr(name,value)
