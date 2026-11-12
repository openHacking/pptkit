#!/usr/bin/env python3
"""Check the named PowerPoint color/bold controls; no claim about arbitrary fonts or hosts."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unpack(p):
 with zipfile.ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
def check(d):
 layout=E.fromstring(d['ppt/slideLayouts/slideLayout3.xml']);assert layout.find('p:cSld',ns).get('name')=='Target'
 items=[s for s in layout.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy'];assert len(items)==1
 level=items[0].find('p:txBody/a:lstStyle/a:lvl1pPr',ns);assert level is not None and not level.attrib and len(level)==1
 r=level.find('a:defRPr',ns);assert r is not None and r.attrib=={'b':'1'} and len(r)==1 and r.find('a:solidFill/a:schemeClr',ns).get('val')=='accent1'
 for i in range(1,4):
  slide=E.fromstring(d[f'ppt/slides/slide{i}.xml']);assert (slide.get('show')=='0')==(i==2)
  rows=[s for s in slide.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')==f'body-{i-1}'];assert len(rows)==1
  run=rows[0].find('p:txBody/a:p/a:r',ns);assert run.find('a:t',ns).text==f'Native reset page {i-1}';prop=run.find('a:rPr',ns);assert prop.find('a:latin',ns).get('typeface')=='Arial' and prop.get('sz')=='2400'
  if i==1:assert 'b' not in prop.attrib and prop.find('a:solidFill',ns) is None
  else:assert prop.get('b')=='0' and prop.find('a:solidFill/a:srgbClr',ns).get('val')=='CC2200'
files=['inherit.pptx','round1.pptx','round2.pptx','round3.pptx'];data={name:unpack(root/name) for name in files}
for x in data.values():check(x)
pdfs=[fitz.open(root/n) for n in ['materialize.pdf','inherit.pdf','round3.pdf']];assert all(len(x)==3 for x in pdfs)
for i in range(3):
 pages=[x[i] for x in pdfs];assert pages[0].get_pixmap(alpha=False).samples==pages[1].get_pixmap(alpha=False).samples==pages[2].get_pixmap(alpha=False).samples
 for method in ('get_texttrace','get_drawings'):
  vals=[getattr(page,method)() for page in pages];assert vals[0]==vals[1]==vals[2]
 vals=[page.get_image_info(hashes=True) for page in pages];assert vals[0]==vals[1]==vals[2]
changed=unpack(root/'layout-edit.pptx');assert all(changed[f'ppt/slides/slide{i}.xml']==data['inherit.pptx'][f'ppt/slides/slide{i}.xml'] for i in range(1,4))
edit=fitz.open(root/'layout-edit.pdf');assert len(edit)==3
for i,page in enumerate(edit):
 spans=[s for b in page.get_text('dict')['blocks'] if 'lines'in b for line in b['lines'] for s in line['spans'] if s['text']==f'Native reset page {i}'];assert len(spans)==1 and spans[0]['font']=='ArialMT' and spans[0]['color']==(0x008800 if i==0 else 0xcc2200)
 if i:assert page.get_pixmap(alpha=False).samples==pdfs[0][i].get_pixmap(alpha=False).samples
negative=root/'negative';negative.mkdir(exist_ok=True);results={}
for name in ['add-font','add-bold','add-color','change-default','hide-first']:
 d=dict(data['round3.pptx']);part='ppt/slideLayouts/slideLayout3.xml' if name=='change-default' else 'ppt/slides/slide1.xml';x=E.fromstring(d[part])
 if name=='change-default':
  s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');s.find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',ns).set('b','0')
 elif name=='hide-first':x.set('show','0')
 else:
  s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='body-0')
  if name=='add-font':
   # Put the forbidden new default in the layout so the scope check catches it.
   part='ppt/slideLayouts/slideLayout3.xml';x=E.fromstring(d[part]);s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');E.SubElement(s.find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',ns),'{'+A+'}latin',{'typeface':'Arial'})
  else:
   r=s.find('p:txBody/a:p/a:r/a:rPr',ns)
   if name=='add-bold':r.set('b','0')
   else:E.SubElement(E.SubElement(r,'{'+A+'}solidFill'),'{'+A+'}srgbClr',{'val':'FF0000'})
 d[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);p=negative/(name+'.pptx')
 with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
  for k,v in d.items():z.writestr(k,v)
 try:check(unpack(p))
 except (AssertionError,AttributeError,KeyError):results[name]={'rejected':True,'sha256':sha(p)}
 else:raise AssertionError('Accepted '+name)
paths=['materialize.pptx','inherit.pptx','round1.pptx','round2.pptx','round3.pptx','materialize.pdf','inherit.pdf','round3.pdf','layout-edit.pptx','layout-edit.pdf'];sources=[Path(__file__),Path('scripts/prepare-native-color-bold.mjs'),Path('tests/fixtures/layout-rebind.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'scope':'one native color/bold-only control, explicit local font; three reopened PPTX, first/final PDFs','directory':str(root),'files':{x:sha(root/x) for x in paths},'sourceSha256':{str(p):sha(p) for p in sources},'pixelsTextTraceDrawingImagesEqual':True,'layoutEditChangedSelectedOnly':True,'negativeZip':results,'generalFontMetricsAccepted':False,'nativeInteractiveEditingVerified':False};Path('docs/testing/native-minimal-color-bold-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
