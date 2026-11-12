#!/usr/bin/env python3
"""Verify empty paragraph defaults have no effect in a bounded native sample."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':A}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unzip(p):
 with zipfile.ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
def check(a,b):
 assert a.keys()==b.keys();assert {n for n in a if a[n]!=b[n]}=={'ppt/slides/slide1.xml','ppt/slides/slide2.xml'}
 for i in (1,2):
  part=f'ppt/slides/slide{i}.xml';old,new=E.fromstring(a[part]),E.fromstring(b[part]);rows=[s for s in new.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('name')==f'body-{i-1}'];assert len(rows)==1;para=rows[0].find('p:txBody/a:p/a:pPr',ns);run=para.find('a:defRPr',ns);assert run is not None and not run.attrib and len(run)==0;para.remove(run);assert E.tostring(old)==E.tostring(new)
a,b=unzip(root/'input.pptx'),unzip(root/'empty-paragraph-default.pptx');check(a,b);negative={}
for i in (1,2):
 bad=dict(b);part=f'ppt/slides/slide{i}.xml';x=E.fromstring(bad[part]);row=next(s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('name')==f'body-{i-1}');row.find('p:txBody/a:p/a:pPr/a:defRPr',ns).set('sz','2400');bad[part]=E.tostring(x)
 try:check(a,bad)
 except AssertionError:negative[str(i)]=True
 else:raise AssertionError(f'accepted nonempty paragraph default slide {i}')
def records(page):return (page.get_pixmap(alpha=False).samples,page.get_texttrace(),page.get_drawings(),page.get_image_info(hashes=True))
a,b,c=[fitz.open(root/n) for n in ('input.pdf','empty-paragraph-default.pdf','default-materialized.pdf')];assert all(len(x)==3 for x in (a,b,c));shifts=[]
for i in range(3):
 assert records(a[i])==records(b[i]);assert (records(b[i])==records(c[i]))==(i==2)
 def body(page):
  rows=[t for t in page.get_texttrace() if ''.join(chr(ch[0]) for ch in t['chars'])==f'Native reset page {i}'];assert len(rows)==1;return rows[0]
 x,y=body(b[i]),body(c[i]);delta=x['bbox'][1]-y['bbox'][1];assert abs(delta-(-0.96 if i<2 else 0))<0.001;shifts.append(delta)
files=['input.pptx','input.pdf','empty-paragraph-default.pptx','empty-paragraph-default.pdf','default-materialized.pdf'];sources=[Path(__file__),Path('scripts/research/prepare-mixed-empty-paragraph-default.py')];report={'scope':'PowerPoint 16.104 empty paragraph default probe, initial open only','directory':str(root),'files':{n:sha(root/n) for n in files},'sourceSha256':{str(p):sha(p) for p in sources},'negativeNonemptyDefaultsRejected':negative,'matchesInput':True,'bodyShiftAgainstMaterializedPoints':shifts,'nativeSaveReopenVerified':False,'productionFixImplemented':False};Path('docs/testing/native-mixed-empty-paragraph-default-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
