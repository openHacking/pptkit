#!/usr/bin/env python3
"""Create independent paragraph mark size probes from a mixed native fixture."""
import sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
r=Path(sys.argv[1]);ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
for source,dest in [('input.pptx','mark-size.pptx'),('mark-effective.pptx','mark-effective-size.pptx')]:
 with zipfile.ZipFile(r/source) as z:parts={n:z.read(n) for n in z.namelist()}
 for i in (1,2):
  part=f'ppt/slides/slide{i}.xml';x=E.fromstring(parts[part]);rows=[s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('name')==f'body-{i-1}'];assert len(rows)==1;mark=rows[0].find('p:txBody/a:p/a:endParaRPr',ns);assert mark is not None and 'sz' not in mark.attrib;mark.set('sz','2400');parts[part]=E.tostring(x,encoding='utf-8',xml_declaration=True)
 with zipfile.ZipFile(r/dest,'w',zipfile.ZIP_DEFLATED) as z:
  for n,b in parts.items():z.writestr(n,b)
