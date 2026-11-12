#!/usr/bin/env python3
"""Apply one independent layout body property change per named control."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from xml.etree import ElementTree as E
import sys
root=Path(sys.argv[1]); output=Path(sys.argv[2]);P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
with ZipFile(root/'inherit.pptx') as z:data={i.filename:z.read(i.filename) for i in z.infolist()}
changes={'wrap':('wrap','square'),'verticalAlign':('anchor','b'),'left':('lIns','635000'),'right':('rIns','635000'),'top':('tIns','635000'),'bottom':('bIns','635000')};seen=[]
for part,blob in list(data.items()):
 if not part.startswith('ppt/slideLayouts/slideLayout') or not part.endswith('.xml'):continue
 x=E.fromstring(blob);key=x.find('p:cSld',ns).get('name').removeprefix('Target ')
 if key not in changes:continue
 matches=[s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('name')=='body'];assert len(matches)==1
 attr,value=changes[key];matches[0].find('p:txBody/a:bodyPr',ns).set(attr,value);data[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);seen.append(key)
assert set(seen)==set(changes)
with ZipFile(output,'w',ZIP_DEFLATED) as z:
 for part,blob in data.items():z.writestr(part,blob)
