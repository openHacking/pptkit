#!/usr/bin/env python3
"""Validate scoped production frame inheritance against actual PowerPoint files."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
attrs={'wrap':'none','anchor':'b','lIns':'381000','tIns':'254000','rIns':'508000','bIns':'127000'}
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def unpack(path):
 with zipfile.ZipFile(path) as z:return {i.filename:z.read(i.filename) for i in z.infolist()}
def shape(xml,name):
 x=E.fromstring(xml);rows=[s for s in x.findall('.//p:sp',ns) if (n:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and n.get('name')==name];assert len(rows)==1;return rows[0]
def check(data):
 layout=E.fromstring(data['ppt/slideLayouts/slideLayout3.xml']);assert layout.find('p:cSld',ns).get('name')=='Target';body=shape(data['ppt/slideLayouts/slideLayout3.xml'],'copy').find('p:txBody/a:bodyPr',ns);assert all(body.get(k)==v for k,v in attrs.items())
 for i in range(1,4):
  part=data[f'ppt/slides/slide{i}.xml'];slide=E.fromstring(part);assert (slide.get('show')=='0')==(i==2);s=shape(part,f'body-{i-1}');b=s.find('p:txBody/a:bodyPr',ns);assert f'Frame native page {i-1}' in ''.join(s.itertext());assert b.find('a:noAutofit',ns) is not None
  if i==1: assert all(k not in b.attrib for k in attrs)
  else: assert b.get('wrap')=='square' and b.get('anchor')=='t' and all(b.get(k)=='25400' for k in ('lIns','rIns','tIns','bIns'))
files=['inherit.pptx','round1.pptx','round2.pptx','round3.pptx'];data={name:unpack(root/name) for name in files}
for d in data.values():check(d)
materialized=unpack(root/'materialize.pptx');baseline=shape(materialized['ppt/slides/slide1.xml'],'body-0').find('p:txBody/a:bodyPr',ns);assert all(baseline.get(k)==v for k,v in attrs.items())
edit=unpack(root/'layout-edit.pptx');original=data['inherit.pptx'];assert edit.keys()==original.keys();assert all(original[name]==edit[name] for name in edit if name!='ppt/slideLayouts/slideLayout3.xml');e=shape(edit['ppt/slideLayouts/slideLayout3.xml'],'copy').find('p:txBody/a:bodyPr',ns);assert e.get('anchor')=='t' and e.get('wrap')=='square' and all(e.get(k)=='25400' for k in ('lIns','rIns','tIns','bIns'))
pdfs=[fitz.open(root/n) for n in ('materialize.pdf','inherit.pdf','round3.pdf','layout-edit.pdf')];assert all(len(x)==3 for x in pdfs)
for i in range(3):
 a,b,c,d=[x[i] for x in pdfs];assert a.get_pixmap(alpha=False).samples==b.get_pixmap(alpha=False).samples==c.get_pixmap(alpha=False).samples
 for method in ('get_texttrace','get_drawings','get_image_info'):
  kw={'hashes':True} if method=='get_image_info' else {};assert getattr(a,method)(**kw)==getattr(b,method)(**kw)==getattr(c,method)(**kw)
 assert (b.get_pixmap(alpha=False).samples==d.get_pixmap(alpha=False).samples)==(i!=0);assert f'Frame native page {i}' in d.get_text()
negative=root/'negative';negative.mkdir(exist_ok=True);results={}
for key in [*attrs,'local','hidden']:
 d=dict(data['round3.pptx']);part='ppt/slideLayouts/slideLayout3.xml' if key in attrs else 'ppt/slides/slide1.xml';x=E.fromstring(d[part])
 if key=='hidden':x.set('show','0')
 else:
  name='copy' if key in attrs else 'body-0'
  for candidate in x.findall('.//p:sp',ns):
   c=candidate.find('p:nvSpPr/p:cNvPr',ns)
   if c is not None and c.get('name')==name: candidate.find('p:txBody/a:bodyPr',ns).set(key if key in attrs else 'lIns','1')
 d[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);path=negative/(key+'.pptx')
 with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
  for n,value in d.items():z.writestr(n,value)
 try:check(unpack(path));raise RuntimeError('negative accepted '+key)
 except AssertionError:results[key]=sha(path)
paths=files+['materialize.pptx','layout-edit.pptx','materialize.pdf','inherit.pdf','round3.pdf','layout-edit.pdf'];sources=[Path(__file__),Path('scripts/prepare-native-frame.mjs'),Path('tests/fixtures/layout-rebind.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'scope':'production six-attribute frame control; three PowerPoint 16.104 SaveAs/close/reopen rounds, first/final PDFs','directory':str(root),'files':{f:sha(root/f) for f in paths},'sourceSha256':{str(p):sha(p) for p in sources},'pixelsTextTraceDrawingImagesEqual':True,'layoutEditChangedSelectedOnly':True,'negativeZip':results,'individualNativePropertyEffectsVerified':False,'nativeInteractiveEditingVerified':False,'autofitVerified':False};Path('docs/testing/native-frame-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
