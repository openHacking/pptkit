#!/usr/bin/env python3
"""Change only the target layout's selected mixed placeholder properties."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from xml.etree import ElementTree as E
import sys
r=Path(sys.argv[1]);P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
with ZipFile(r/'input.pptx') as z:parts={n:z.read(n) for n in z.namelist()}
part='ppt/slideLayouts/slideLayout3.xml';x=E.fromstring(parts[part]);s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');s.find('p:txBody/a:bodyPr',ns).set('wrap','none');p=s.find('p:txBody/a:lstStyle/a:lvl1pPr',ns);p.set('algn','r');run=p.find('a:defRPr',ns);run.set('b','0')
for tag in ('latin','ea','cs'):run.find('a:'+tag,ns).set('typeface','Georgia')
fill=run.find('a:solidFill',ns);fill.clear();E.SubElement(fill,f'{{{A}}}srgbClr',{'val':'008800'})
parts[part]=E.tostring(x,encoding='utf-8',xml_declaration=True)
with ZipFile(r/'mixed-layout-edit.pptx','w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
