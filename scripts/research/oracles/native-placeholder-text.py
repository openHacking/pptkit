#!/usr/bin/env python3
"""Bounded independent checks for opt-in native run inheritance and shared-layout regressions."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import fitz
P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
root=Path(sys.argv[1]).resolve();historical=Path(sys.argv[2]).resolve()
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def archive(path):
 with zipfile.ZipFile(path) as z:return {n:z.read(n) for n in z.namelist()}
def selected(entries,slide_i):
 x=ET.fromstring(entries[f'ppt/slides/slide{slide_i}.xml']); assert (x.get('show')=='0')==(slide_i==2)
 rows=[e for e in x.findall('.//p:sp',ns) if e.find('p:nvSpPr/p:cNvPr',ns) is not None and e.find('p:nvSpPr/p:cNvPr',ns).get('name')==f'body-{slide_i-1}'];assert len(rows)==1
 return rows[0]
def check(entries):
 layouts=[]
 for name,data in entries.items():
  if name.startswith('ppt/slideLayouts/slideLayout') and name.endswith('.xml'):
   x=ET.fromstring(data);c=x.find('p:cSld',ns)
   if c is not None and c.get('name')=='Target':layouts.append(x)
 assert len(layouts)==1
 shapes=[x for x in layouts[0].findall('.//p:sp',ns) if x.find('p:nvSpPr/p:cNvPr',ns) is not None and x.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy'];assert len(shapes)==1
 default=shapes[0].find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',ns);assert default is not None and default.get('b')=='1' and default.find('a:latin',ns).get('typeface')=='+mn-lt' and default.find('a:solidFill/a:schemeClr',ns).get('val')=='accent1'
 for i in range(1,4):
  x=selected(entries,i);run=x.find('p:txBody/a:p/a:r',ns);assert run.find('a:t',ns).text==f'Native reset page {i-1}';r=run.find('a:rPr',ns);assert r.get('sz')=='2400'
  if i==1:
   assert 'b' not in r.attrib and all(r.find('a:'+tag,ns) is None for tag in ('solidFill','latin','ea','cs'))
  else:
   assert r.get('b')=='0' and r.find('a:latin',ns).get('typeface')=='Arial' and r.find('a:solidFill/a:srgbClr',ns).get('val')=='CC2200'
for name in ('input.pptx','round1.pptx','round2.pptx','round3.pptx'):check(archive(root/name))
assert (root/'default-materialized.pptx').read_bytes()==(historical/'input.pptx').read_bytes()
def extract(path):
 doc=fitz.open(path);assert len(doc)==3;rows=[]
 for i,page in enumerate(doc):
  spans=[s for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['text']==f'Native reset page {i}'];assert len(spans)==1;s=spans[0];assert s['font']==('CourierNewPS-BoldMT' if i==0 else 'ArialMT') and s['color']==(0x0044cc if i==0 else 0xcc2200);rows.append({k:s[k] for k in ('text','font','size','color','bbox')})
 return rows
first=extract(root/'first.pdf');last=extract(root/'round3.pdf');old=extract(historical/'round3.pdf');assert first==last
for i in range(3):
 a=fitz.open(root/'first.pdf')[i].get_pixmap(alpha=False);b=fitz.open(root/'round3.pdf')[i].get_pixmap(alpha=False);assert a.samples==b.samples
edited=fitz.open(root/'layout-edit.pdf'); assert len(edited)==3
for i,page in enumerate(edited):
 spans=[s for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['text']==f'Native reset page {i}']; assert len(spans)==1
 assert spans[0]['font']==('Georgia' if i==0 else 'ArialMT') and spans[0]['color']==(0x008800 if i==0 else 0xcc2200)
source=archive(root/'input.pptx'); changed=archive(root/'layout-edit.pptx'); assert all(source[f'ppt/slides/slide{i}.xml']==changed[f'ppt/slides/slide{i}.xml'] for i in range(1,4))
changed_layout=ET.fromstring(changed['ppt/slideLayouts/slideLayout3.xml']); shape=next(x for x in changed_layout.findall('.//p:sp',ns) if x.find('p:nvSpPr/p:cNvPr',ns) is not None and x.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');rp=shape.find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',ns);assert rp.get('b')=='0' and rp.find('a:latin',ns).get('typeface')=='Georgia' and rp.find('a:solidFill/a:srgbClr',ns).get('val')=='008800'
# The control on page 1 shares the modified target layout and moves despite identical local declarations.
dy=[first[i]['bbox'][1]-old[i]['bbox'][1] for i in range(3)];assert abs(dy[1])>0.5 and abs(dy[2])<0.01
negdir=root/'negatives';negdir.mkdir(exist_ok=True);base=archive(root/'round3.pptx');negative={}
for name in ('forced-bold','forced-font','forced-color','target-default','hidden'):
 data=dict(base);part='ppt/slideLayouts/slideLayout3.xml' if name=='target-default' else 'ppt/slides/slide1.xml';x=ET.fromstring(data[part]);
 if name=='target-default':
  shape=next(y for y in x.findall('.//p:sp',ns) if y.find('p:nvSpPr/p:cNvPr',ns) is not None and y.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');shape.find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',ns).set('b','0')
 elif name=='hidden':x.set('show','0')
 else:
  shape=next(y for y in x.findall('.//p:sp',ns) if y.find('p:nvSpPr/p:cNvPr',ns) is not None and y.find('p:nvSpPr/p:cNvPr',ns).get('name')=='body-0');r=shape.find('p:txBody/a:p/a:r/a:rPr',ns)
  if name=='forced-bold':r.set('b','0')
  elif name=='forced-font':ET.SubElement(r,'{'+A+'}latin',{'typeface':'Arial'})
  else:ET.SubElement(ET.SubElement(r,'{'+A+'}solidFill'),'{'+A+'}srgbClr',{'val':'000000'})
 data[part]=ET.tostring(x,encoding='utf-8',xml_declaration=True);path=negdir/(name+'.pptx')
 with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
  for key,value in data.items():z.writestr(key,value)
 try:check(archive(path))
 except (AssertionError,AttributeError,KeyError):negative[name]={'sha256':sha(path),'rejected':True}
 else:raise AssertionError('Corrupt file accepted: '+name)
paths=['default-materialized.pptx','input.pptx','round1.pptx','round2.pptx','round3.pptx','first.pdf','round3.pdf','mark-probe.pptx','mark-probe.pdf','mark2-probe.pptx','mark2-probe.pdf','layout-edit.pptx','layout-edit.pdf'];sources=[Path(__file__),Path('scripts/prepare-style-reset-native.mjs'),Path('tests/fixtures/layout-rebind.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'scope':'Opt-in fontFamily/color/bold native inheritance; actual first and third-round PDF; three saved/reopened packages','directory':str(root),'files':{x:sha(root/x) for x in paths},'historicalBaseline':{'file':str(historical/'round3.pdf'),'sha256':sha(historical/'round3.pdf')},'sourceSha256':{str(x):sha(x) for x in sources},'firstAndThirdPixelIdentical':True,'spanRecords':first,'baselineVerticalShiftPoints':dy,'defaultMatchesHistoricalPptxByteForByte':True,'sharedMetricsWarningRequired':True,'markStyleProbesFixedShift':False,'negativeZip':negative,'layoutXmlEditObserved':True,'allPropertiesAccepted':False,'nativeInteractiveLayoutEditingVerified':False}
Path('docs/testing/native-placeholder-text-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
