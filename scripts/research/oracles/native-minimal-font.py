#!/usr/bin/env python3
"""Independent checks of the named three-round native font/color/bold control."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();prior=Path(sys.argv[2]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unzip(p):
 with zipfile.ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
def check(d):
 x=E.fromstring(d['ppt/slideLayouts/slideLayout3.xml']);assert x.find('p:cSld',ns).get('name')=='Target';s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');lvl=s.find('p:txBody/a:lstStyle/a:lvl1pPr',ns);assert not lvl.attrib and len(lvl)==1;r=lvl.find('a:defRPr',ns);assert r.attrib=={'b':'1'} and r.find('a:solidFill/a:schemeClr',ns).get('val')=='accent1'
 for tag in ('latin','ea','cs'):assert r.find('a:'+tag,ns).get('typeface')=='+mn-lt'
 assert len(r)==4
 for i in range(1,4):
  x=E.fromstring(d[f'ppt/slides/slide{i}.xml']);assert (x.get('show')=='0')==(i==2);s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')==f'body-{i-1}');run=s.find('p:txBody/a:p/a:r',ns);assert run.find('a:t',ns).text==f'Native reset page {i-1}';r=run.find('a:rPr',ns);assert r.get('sz')=='2400'
  if i==1:assert 'b' not in r.attrib and all(r.find('a:'+t,ns) is None for t in ('solidFill','latin','ea','cs'))
  else:assert r.get('b')=='0' and r.find('a:latin',ns).get('typeface')=='Arial' and r.find('a:solidFill/a:srgbClr',ns).get('val')=='CC2200'
assert (root/'default-materialized.pptx').read_bytes()==(prior/'default-materialized.pptx').read_bytes()
files=['input.pptx','round1.pptx','round2.pptx','round3.pptx'];data={n:unzip(root/n) for n in files}
for x in data.values():check(x)
a,b,c=fitz.open(root/'first.pdf'),fitz.open(root/'round3.pdf'),fitz.open(prior/'first.pdf');assert len(a)==len(b)==len(c)==3
for i in range(3):
 pages=[a[i],b[i],c[i]];assert pages[0].get_pixmap(alpha=False).samples==pages[1].get_pixmap(alpha=False).samples==pages[2].get_pixmap(alpha=False).samples
 for method in ('get_texttrace','get_drawings'):
  v=[getattr(p,method)() for p in pages];assert v[0]==v[1]==v[2]
 v=[p.get_image_info(hashes=True) for p in pages];assert v[0]==v[1]==v[2]
edit=fitz.open(root/'layout-edit.pdf');changed=unzip(root/'layout-edit.pptx');assert len(edit)==3 and all(data['input.pptx'][f'ppt/slides/slide{i}.xml']==changed[f'ppt/slides/slide{i}.xml'] for i in range(1,4))
for i,p in enumerate(edit):
 spans=[s for block in p.get_text('dict')['blocks'] if 'lines'in block for l in block['lines'] for s in l['spans'] if s['text']==f'Native reset page {i}'];assert len(spans)==1 and spans[0]['font']==('Georgia' if i==0 else 'ArialMT') and spans[0]['color']==(0x008800 if i==0 else 0xcc2200)
negative=root/'negative';negative.mkdir(exist_ok=True);results={}
for name in ('forced-font','forced-color','forced-bold','changed-layout','hidden'):
 d=dict(data['round3.pptx']);part='ppt/slideLayouts/slideLayout3.xml' if name=='changed-layout' else 'ppt/slides/slide1.xml';x=E.fromstring(d[part])
 if name=='hidden':x.set('show','0')
 elif name=='changed-layout':
  s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');s.find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',ns).set('b','0')
 else:
  s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='body-0');r=s.find('p:txBody/a:p/a:r/a:rPr',ns)
  if name=='forced-bold':r.set('b','0')
  elif name=='forced-font':E.SubElement(r,'{'+A+'}latin',{'typeface':'Arial'})
  else:E.SubElement(E.SubElement(r,'{'+A+'}solidFill'),'{'+A+'}srgbClr',{'val':'000000'})
 d[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);p=negative/(name+'.pptx')
 with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
  for k,v in d.items():z.writestr(k,v)
 try:check(unzip(p))
 except (AssertionError,KeyError,AttributeError,StopIteration):results[name]={'sha256':sha(p),'rejected':True}
 else:raise AssertionError('Accepted '+name)
paths=files+['first.pdf','round3.pdf','layout-edit.pptx','layout-edit.pdf','default-materialized.pptx'];src=[Path(__file__),Path('scripts/prepare-style-reset-native.mjs'),Path('tests/fixtures/layout-rebind.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'scope':'named minimal font/color/bold profile; PowerPoint three saved/reopened PPTX, first and third PDF','directory':str(root),'files':{p:sha(root/p) for p in paths},'sourceSha256':{str(p):sha(p) for p in src},'priorOptInBaseline':{'file':str(prior/'first.pdf'),'sha256':sha(prior/'first.pdf')},'threePdfPixelGlyphDrawingImageEquality':True,'layoutEditChangedSelectedOnly':True,'negativeZip':results,'defaultPptxMatchesPrior':True,'sharedFontMetricsShiftResolved':False,'nativeInteractiveEditingVerified':False};Path('docs/testing/native-minimal-font-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
