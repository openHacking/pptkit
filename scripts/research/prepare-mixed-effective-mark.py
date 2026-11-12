#!/usr/bin/env python3
"""Patch only two paragraph marks in a mixed-inheritance research copy."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from xml.etree import ElementTree as E
import sys
r=Path(sys.argv[1]);P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
with ZipFile(r/'input.pptx') as z:parts={n:z.read(n) for n in z.namelist()}
for i,font in ((1,'Courier New'),(2,'Arial')):
 part=f'ppt/slides/slide{i}.xml';x=E.fromstring(parts[part]);sp=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')==f'body-{i-1}');mark=sp.find('p:txBody/a:p/a:endParaRPr',ns)
 for slot in ('latin','ea','cs'):E.SubElement(mark,f'{{{A}}}{slot}',{'typeface':font})
 parts[part]=E.tostring(x,encoding='utf-8',xml_declaration=True)
with ZipFile(r/'mark-effective.pptx','w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)

# Second counterexample changes only the selected paragraph mark's bold flag.
part='ppt/slides/slide1.xml';x=E.fromstring(parts[part]);sp=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='body-0');sp.find('p:txBody/a:p/a:endParaRPr',ns).set('b','1');parts[part]=E.tostring(x,encoding='utf-8',xml_declaration=True)
with ZipFile(r/'mark-effective-bold.pptx','w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
