#!/usr/bin/env python3
"""Verify bounded PowerPoint paragraph-mark counterexample."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
r=Path(sys.argv[1]);A='http://schemas.openxmlformats.org/drawingml/2006/main';P='http://schemas.openxmlformats.org/presentationml/2006/main';ns={'a':A,'p':P}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unzip(p):
 with zipfile.ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
a,b=unzip(r/'input.pptx'),unzip(r/'mark-effective.pptx');assert a.keys()==b.keys()
assert {n for n in a if a[n]!=b[n]}=={'ppt/slides/slide1.xml','ppt/slides/slide2.xml'}
for i,font in ((1,'Courier New'),(2,'Arial')):
 part=f'ppt/slides/slide{i}.xml';old,new=E.fromstring(a[part]),E.fromstring(b[part]);sp=next(s for s in new.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')==f'body-{i-1}');mark=sp.find('p:txBody/a:p/a:endParaRPr',ns)
 for tag in ('latin','ea','cs'):
  child=mark.find('a:'+tag,ns);assert child is not None and child.attrib=={'typeface':font};mark.remove(child)
 assert E.tostring(old)==E.tostring(new)
bold=unzip(r/'mark-effective-bold.pptx');assert bold.keys()==b.keys();assert {n for n in b if b[n]!=bold[n]}=={'ppt/slides/slide1.xml'}
old,new=E.fromstring(b['ppt/slides/slide1.xml']),E.fromstring(bold['ppt/slides/slide1.xml']);sp=next(s for s in new.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='body-0');mark=sp.find('p:txBody/a:p/a:endParaRPr',ns);assert mark.attrib.pop('b')=='1';assert E.tostring(old)==E.tostring(new)
p,q,bb=fitz.open(r/'mark-effective.pdf'),fitz.open(r/'default-materialized.pdf'),fitz.open(r/'mark-effective-bold.pdf');assert len(p)==len(q)==len(bb)==3
for x,y in zip(p,bb):assert x.get_pixmap().samples==y.get_pixmap().samples and x.get_texttrace()==y.get_texttrace() and x.get_drawings()==y.get_drawings() and x.get_image_info(hashes=True)==y.get_image_info(hashes=True)
shifts=[]
for i,(x,y) in enumerate(zip(p,q)):
 assert (x.get_pixmap().samples==y.get_pixmap().samples)==(i!=0);assert (x.get_texttrace()==y.get_texttrace())==(i!=0);assert x.get_drawings()==y.get_drawings();assert x.get_image_info(hashes=True)==y.get_image_info(hashes=True)
 def body(page):
  rows=[t for t in page.get_texttrace() if ''.join(chr(c[0]) for c in t['chars'])==f'Native reset page {i}'];assert len(rows)==1;return rows[0]
 a1,b1=body(x),body(y);assert a1['font']==b1['font'];delta=a1['bbox'][1]-b1['bbox'][1];assert abs(delta-(-0.96 if i==0 else 0))<0.001;shifts.append(delta)
paths=['input.pptx','mark-effective.pptx','mark-effective.pdf','mark-effective-bold.pptx','mark-effective-bold.pdf','default-materialized.pptx','default-materialized.pdf'];sources=[Path(__file__),Path('scripts/research/prepare-mixed-effective-mark.py')]
report={'scope':'manual effective paragraph-mark font probe, PowerPoint 16.104 initial render only','directory':str(r.resolve()),'files':{n:sha(r/n) for n in paths},'sourceSha256':{str(s):sha(s) for s in sources},'bodyShiftPoints':shifts,'selectedPageStillShifted':True,'markBoldDidNotChangeResult':True,'localControlRestored':True,'nativeSaveReopenVerified':False,'productionFixImplemented':False};Path('docs/testing/native-mixed-effective-mark-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
