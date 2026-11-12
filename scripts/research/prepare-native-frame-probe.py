#!/usr/bin/env python3
"""Build manually patched controls for testing PowerPoint bodyPr inheritance."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as E
from tempfile import mkdtemp
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'; ns={'p':P,'a':A}
source=Path('.tmp/native-paragraph/run-XPcnYn/materialize.pptx')
out=Path(mkdtemp(prefix='run-',dir='.tmp/native-frame-probe')).resolve()
with ZipFile(source) as z: base={i.filename:z.read(i.filename) for i in z.infolist()}
def shape(root,name):
 matches=[s for s in root.findall('.//p:sp',ns) if (n:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and n.get('name')==name]
 assert len(matches)==1;return matches[0]
props={'wrap':'none','anchor':'b','lIns':'381000','tIns':'254000','rIns':'508000','bIns':'127000'}
for mode in ('baseline','inherit','edit'):
 data=base.copy();part='ppt/slideLayouts/slideLayout3.xml';root=E.fromstring(data[part]);body=shape(root,'copy').find('p:txBody/a:bodyPr',ns);body.attrib.update(props)
 if mode=='edit': body.attrib.update({'wrap':'square','anchor':'t','lIns':'25400','tIns':'25400','rIns':'25400','bIns':'25400'})
 data[part]=E.tostring(root,encoding='utf-8',xml_declaration=True)
 part='ppt/slides/slide1.xml';root=E.fromstring(data[part]);body=shape(root,'body-0').find('p:txBody/a:bodyPr',ns)
 if mode=='baseline':body.attrib.update(props)
 else:
  for key in props:body.attrib.pop(key,None)
 data[part]=E.tostring(root,encoding='utf-8',xml_declaration=True)
 with ZipFile(out/(mode+'.pptx'),'w',ZIP_DEFLATED) as z:
  for name,content in data.items():z.writestr(name,content)
print(out)
