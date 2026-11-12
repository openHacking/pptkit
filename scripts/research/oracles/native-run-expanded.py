#!/usr/bin/env python3
"""Independent native control for five selectively inherited run properties."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A};attrs={'lang':'fr-FR','sz':'2600','i':'1','u':'sng','strike':'sngStrike'}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unzip(p):
 with zipfile.ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
def check(d):
 x=E.fromstring(d['ppt/slideLayouts/slideLayout3.xml']);assert x.find('p:cSld',ns).get('name')=='Target';s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');level=s.find('p:txBody/a:lstStyle/a:lvl1pPr',ns);assert not level.attrib and len(level)==1;r=level.find('a:defRPr',ns);assert r.attrib==attrs and len(r)==0
 for i in range(1,4):
  x=E.fromstring(d[f'ppt/slides/slide{i}.xml']);assert (x.get('show')=='0')==(i==2);s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')==f'body-{i-1}');run=s.find('p:txBody/a:p/a:r',ns);assert run.find('a:t',ns).text==f'Expanded native page {i-1}';r=run.find('a:rPr',ns);assert r.find('a:latin',ns).get('typeface')=='Arial' and r.find('a:solidFill/a:srgbClr',ns).get('val')=='CC2200'
  if i==1:assert all(key not in r.attrib for key in attrs)
  else:assert all(r.get(key)==value for key,value in {'lang':'en-US','sz':'2000','i':'0','u':'none','strike':'noStrike'}.items())
files=['inherit.pptx','round1.pptx','round2.pptx','round3.pptx'];data={f:unzip(root/f) for f in files}
for d in data.values():check(d)
pdfs=[fitz.open(root/n) for n in ('materialize.pdf','inherit.pdf','round3.pdf')];assert all(len(d)==3 for d in pdfs)
for i in range(3):
 pages=[d[i] for d in pdfs];assert pages[0].get_pixmap(alpha=False).samples==pages[1].get_pixmap(alpha=False).samples==pages[2].get_pixmap(alpha=False).samples
 for method in ('get_texttrace','get_drawings'):
  v=[getattr(p,method)() for p in pages];assert v[0]==v[1]==v[2]
 v=[p.get_image_info(hashes=True) for p in pages];assert v[0]==v[1]==v[2]
edit=unzip(root/'layout-edit.pptx');assert all(edit[f'ppt/slides/slide{i}.xml']==data['inherit.pptx'][f'ppt/slides/slide{i}.xml'] for i in range(1,4));x=E.fromstring(edit['ppt/slideLayouts/slideLayout3.xml']);s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');r=s.find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',ns);assert r.attrib=={'lang':'en-US','sz':'2000','i':'0','u':'none','strike':'noStrike'}
changed=fitz.open(root/'layout-edit.pdf');assert len(changed)==3
for i,p in enumerate(changed):
 spans=[s for b in p.get_text('dict')['blocks'] if 'lines'in b for l in b['lines'] for s in l['spans'] if s['text']==f'Expanded native page {i}'];assert len(spans)==1
 if i==0:assert spans[0]['font']=='Arial-BoldMT' and 20<spans[0]['size']<22
 else:assert p.get_pixmap(alpha=False).samples==pdfs[0][i].get_pixmap(alpha=False).samples
# Two red decoration rectangles appear in the target text band, and disappear after layout edit.
def decoration(p):return [d for d in p.get_drawings() if d['type']=='f' and d.get('fill') is not None and abs(d['fill'][0]-.8)<.01 and abs(d['fill'][1]-.1333)<.01 and 140<d['rect'].y0<180]
assert len(decoration(pdfs[1][0]))==2 and len(decoration(changed[0]))==0
negative=root/'negative';negative.mkdir(exist_ok=True);results={}
for name in ['lang','sz','i','u','strike','layout','hidden']:
 d=dict(data['round3.pptx']);part='ppt/slideLayouts/slideLayout3.xml' if name=='layout' else 'ppt/slides/slide1.xml';x=E.fromstring(d[part])
 if name=='hidden':x.set('show','0')
 elif name=='layout':
  s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');s.find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',ns).set('sz','2000')
 else:
  s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='body-0');s.find('p:txBody/a:p/a:r/a:rPr',ns).set(name,{'lang':'en-US','sz':'2000','i':'0','u':'none','strike':'noStrike'}[name])
 d[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);p=negative/(name+'.pptx')
 with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
  for k,v in d.items():z.writestr(k,v)
 try:check(unzip(p))
 except (AssertionError,AttributeError,KeyError,StopIteration):results[name]={'rejected':True,'sha256':sha(p)}
 else:raise AssertionError('Accepted '+name)
paths=files+['materialize.pptx','layout-edit.pptx','materialize.pdf','inherit.pdf','round3.pdf','layout-edit.pdf'];sources=[Path(__file__),Path('scripts/prepare-native-run-expanded.mjs'),Path('tests/fixtures/layout-rebind.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'scope':'named fontSize/italic/underline/strike/language control, three saved and reopened PPTX, first/final PDF','directory':str(root),'files':{x:sha(root/x) for x in paths},'sourceSha256':{str(p):sha(p) for p in sources},'pixelsTextTraceDrawingImagesEqual':True,'layoutEditChangedSelectedOnly':True,'decorationRectanglesBefore':2,'decorationRectanglesAfter':0,'negativeZip':results,'nativeInteractiveEditingVerified':False,'sharedFontMetricsResolved':False};Path('docs/testing/native-run-expanded-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
