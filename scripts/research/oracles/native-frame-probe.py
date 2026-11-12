#!/usr/bin/env python3
"""Check actual PowerPoint PDF observations for a manually patched frame probe."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E
import hashlib,json,sys,fitz
root=Path(sys.argv[1]).resolve(); P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
def unpack(name):
 with ZipFile(root/name) as z:return {p:z.read(p) for p in z.namelist()}
def shape(blob,name):
 r=E.fromstring(blob);x=[s for s in r.findall('.//p:sp',ns) if (n:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and n.get('name')==name];assert len(x)==1;return x[0]
base,inherit,edit=[unpack(n+'.pptx') for n in ('baseline','inherit','edit')];slide='ppt/slides/slide1.xml';layout='ppt/slideLayouts/slideLayout3.xml';attrs={'wrap':'none','anchor':'b','lIns':'381000','tIns':'254000','rIns':'508000','bIns':'127000'}
for d in (base,inherit):
 p=shape(d[layout],'copy').find('p:txBody/a:bodyPr',ns);assert all(p.get(k)==v for k,v in attrs.items())
b=shape(base[slide],'body-0').find('p:txBody/a:bodyPr',ns);i=shape(inherit[slide],'body-0').find('p:txBody/a:bodyPr',ns);assert all(b.get(k)==v and k not in i.attrib for k,v in attrs.items());assert inherit[slide]==edit[slide]
changed=shape(edit[layout],'copy').find('p:txBody/a:bodyPr',ns);assert changed.get('anchor')=='t' and changed.get('lIns')=='25400'
pdfs=[fitz.open(root/(n+'.pdf')) for n in ('baseline','inherit','edit')];assert all(len(x)==3 for x in pdfs)
for index in range(3):
 a,b,c=(d[index] for d in pdfs);assert a.get_pixmap(alpha=False).samples==b.get_pixmap(alpha=False).samples;assert a.get_texttrace()==b.get_texttrace();assert a.get_drawings()==b.get_drawings();assert a.get_image_info(hashes=True)==b.get_image_info(hashes=True)
 assert (b.get_pixmap(alpha=False).samples==c.get_pixmap(alpha=False).samples)==(index!=0)
 assert f'Paragraph native page {index}' in c.get_text()
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
files=[n+'.'+ext for n in ('baseline','inherit','edit') for ext in ('pptx','pdf')];report={'scope':'manual six-attribute frame probe; three PowerPoint 16.104 first-open PDFs; combined effects only','directory':str(root),'files':{n:sha(root/n) for n in files},'producer':{'path':'scripts/research/prepare-native-frame-probe.py','sha256':sha(Path('scripts/research/prepare-native-frame-probe.py'))},'baselineEqualsInheritedPixelsTextDrawingsImages':True,'layoutOnlyChangeAffectsSelectedPage':True,'individualPropertiesProven':False,'productionOutputProven':False,'saveReopenProven':False,'interactiveLayoutEditProven':False}
Path('docs/testing/native-frame-probe-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
