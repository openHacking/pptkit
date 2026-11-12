#!/usr/bin/env python3
"""Add an empty paragraph default without setting font or run values."""
import sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
r=Path(sys.argv[1]);A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':A}
with zipfile.ZipFile(r/'input.pptx') as z:parts={n:z.read(n) for n in z.namelist()}
for i in (1,2):
 part=f'ppt/slides/slide{i}.xml';x=E.fromstring(parts[part]);rows=[s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('name')==f'body-{i-1}'];assert len(rows)==1;para=rows[0].find('p:txBody/a:p/a:pPr',ns);assert para is not None and para.find('a:defRPr',ns) is None;E.SubElement(para,f'{{{A}}}defRPr');parts[part]=E.tostring(x,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(r/'empty-paragraph-default.pptx','w',zipfile.ZIP_DEFLATED) as z:
 for n,b in parts.items():z.writestr(n,b)
