#!/usr/bin/env python3
"""Verify a bounded PowerPoint no-paragraph-mark counterexample."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unzip(p):
 with zipfile.ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
def check(a,b):
 assert a.keys()==b.keys();assert {n for n in a if a[n]!=b[n]}=={'ppt/slides/slide1.xml','ppt/slides/slide2.xml'}
 for i in (1,2):
  part=f'ppt/slides/slide{i}.xml';old,new=E.fromstring(a[part]),E.fromstring(b[part]);rows=[s for s in new.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('name')==f'body-{i-1}'];assert len(rows)==1;para=rows[0].find('p:txBody/a:p',ns);assert para.find('a:endParaRPr',ns) is None;E.SubElement(para,'{'+ns['a']+'}endParaRPr',{'lang':'en-US'});assert E.tostring(old)==E.tostring(new)
a,b=unzip(root/'input.pptx'),unzip(root/'no-mark.pptx');check(a,b);negative={}
for i in (1,2):
 bad=dict(b);part=f'ppt/slides/slide{i}.xml';x=E.fromstring(bad[part]);row=next(s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('name')==f'body-{i-1}');E.SubElement(row.find('p:txBody/a:p',ns),'{'+ns['a']+'}endParaRPr',{'lang':'en-US'});bad[part]=E.tostring(x)
 try:check(a,bad)
 except AssertionError:negative[str(i)]=True
 else:raise AssertionError(f'accepted restored mark slide {i}')
def records(page):return (page.get_pixmap(alpha=False).samples,page.get_texttrace(),page.get_drawings(),page.get_image_info(hashes=True))
a,b,c=[fitz.open(root/n) for n in ('no-mark.pdf','mark-effective.pdf','default-materialized.pdf')];assert all(len(x)==3 for x in (a,b,c));shifts=[]
for i in range(3):
 assert records(a[i])==records(b[i]);assert (records(a[i])==records(c[i]))==(i!=0)
 def body(page):
  rows=[t for t in page.get_texttrace() if ''.join(chr(ch[0]) for ch in t['chars'])==f'Native reset page {i}'];assert len(rows)==1;return rows[0]
 x,y=body(a[i]),body(c[i]);delta=x['bbox'][1]-y['bbox'][1];assert abs(delta-(-0.96 if i==0 else 0))<0.001;shifts.append(delta)
files=['input.pptx','no-mark.pptx','no-mark.pdf','mark-effective.pdf','default-materialized.pdf'];sources=[Path(__file__),Path('scripts/research/prepare-mixed-no-mark.py')];report={'scope':'PowerPoint 16.104 missing paragraph mark probe, initial open only','directory':str(root),'files':{n:sha(root/n) for n in files},'sourceSha256':{str(p):sha(p) for p in sources},'negativeRestoredMarksRejected':negative,'matchesEffectiveFontMarkProbe':True,'bodyShiftAgainstMaterializedPoints':shifts,'nativeSaveReopenVerified':False,'productionFixImplemented':False};Path('docs/testing/native-mixed-no-mark-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
