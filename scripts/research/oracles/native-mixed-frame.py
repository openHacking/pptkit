#!/usr/bin/env python3
"""Independent checks of the named three-round native font/color/bold control."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unzip(p):
 with zipfile.ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
def check(d):
 x=E.fromstring(d['ppt/slideLayouts/slideLayout3.xml']);assert x.find('p:cSld',ns).get('name')=='Target';s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');lvl=s.find('p:txBody/a:lstStyle/a:lvl1pPr',ns);assert lvl.attrib=={'algn':'l'} and len(lvl)==1;r=lvl.find('a:defRPr',ns);assert r.attrib=={'b':'1'} and r.find('a:solidFill/a:schemeClr',ns).get('val')=='accent1'
 for tag in ('latin','ea','cs'):assert r.find('a:'+tag,ns).get('typeface')=='+mn-lt'
 assert len(r)==4
 for i in range(1,4):
  x=E.fromstring(d[f'ppt/slides/slide{i}.xml']);assert (x.get('show')=='0')==(i==2);s=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')==f'body-{i-1}');run=s.find('p:txBody/a:p/a:r',ns);assert run.find('a:t',ns).text==f'Native reset page {i-1}';r=run.find('a:rPr',ns);assert r.get('sz')=='2400'
  body=s.find('p:txBody/a:bodyPr',ns)
  if i==1:
   assert 'wrap' not in body.attrib
   assert 'b' not in r.attrib and all(r.find('a:'+t,ns) is None for t in ('solidFill','latin','ea','cs')); assert 'algn' not in s.find('p:txBody/a:p/a:pPr',ns).attrib
  else:assert body.get('wrap')=='none' and r.get('b')=='0' and r.find('a:latin',ns).get('typeface')=='Arial' and r.find('a:solidFill/a:srgbClr',ns).get('val')=='CC2200'

files=['input.pptx','round1.pptx','round2.pptx','round3.pptx'];data={n:unzip(root/n) for n in files}
for d in data.values():check(d)
a,b,c=[fitz.open(root/n) for n in ('input.pdf','round3.pdf','default-materialized.pdf')];assert len(a)==len(b)==len(c)==3
shifts=[]
for i in range(3):
 x,y,z=a[i],b[i],c[i];assert x.get_pixmap(alpha=False).samples==y.get_pixmap(alpha=False).samples;assert x.get_texttrace()==y.get_texttrace();assert x.get_drawings()==y.get_drawings();assert x.get_image_info(hashes=True)==y.get_image_info(hashes=True)
 assert (x.get_pixmap(alpha=False).samples==z.get_pixmap(alpha=False).samples)==(i==2)
 def find(p):
  rows=[t for t in p.get_texttrace() if ''.join(chr(ch[0]) for ch in t['chars'])==f'Native reset page {i}'];assert len(rows)==1;return rows[0]
 q,v=find(x),find(z);assert q['font']==v['font'] and q['size']==v['size'];delta=q['bbox'][1]-v['bbox'][1];assert abs(delta-(-0.96 if i<2 else 0))<0.001;shifts.append(delta)
# Package-only layout changes must preserve every other part, including slide XML.
changed=unzip(root/'mixed-layout-edit.pptx');base=data['input.pptx'];assert changed.keys()==base.keys();assert {n for n in changed if changed[n]!=base[n]}=={'ppt/slideLayouts/slideLayout3.xml'}
part='ppt/slideLayouts/slideLayout3.xml';original=E.fromstring(base[part]);edited=E.fromstring(changed[part]);sp=next(s for s in edited.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy');body=sp.find('p:txBody/a:bodyPr',ns);assert body.get('wrap')=='none';body.set('wrap','square');lvl=sp.find('p:txBody/a:lstStyle/a:lvl1pPr',ns);assert lvl.get('algn')=='r';lvl.set('algn','l');rr=lvl.find('a:defRPr',ns);assert rr.get('b')=='0';rr.set('b','1')
for tag in ('latin','ea','cs'):
 slot=rr.find('a:'+tag,ns);assert slot.get('typeface')=='Georgia';slot.set('typeface','+mn-lt')
fill=rr.find('a:solidFill',ns);assert fill.find('a:srgbClr',ns).get('val')=='008800';fill.clear();E.SubElement(fill,'{'+A+'}schemeClr',{'val':'accent1'});assert E.tostring(original)==E.tostring(edited)
edit=fitz.open(root/'mixed-layout-edit.pdf');assert len(edit)==3
for i,(initial,changed_page) in enumerate(zip(a,edit)):
 assert (initial.get_pixmap().samples==changed_page.get_pixmap().samples)==(i==2)
 def span(page):
  rows=[s for block in page.get_text('dict')['blocks'] if 'lines' in block for line in block['lines'] for s in line['spans'] if s['text']==f'Native reset page {i}'];assert len(rows)==1;return rows[0]
 old_span,new_span=span(initial),span(changed_page)
 assert new_span['font']==('Georgia' if i==0 else 'ArialMT') and new_span['color']==(0x008800 if i==0 else 0xcc2200)
 if i==1:assert abs((new_span['bbox'][1]-old_span['bbox'][1])-0.96)<0.001
negative={}
for name in ('forced-wrap','forced-bold','forced-font','forced-align','hidden'):
 d=dict(data['round3.pptx']);part='ppt/slides/slide1.xml';x=E.fromstring(d[part]);sp=next(s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='body-0')
 if name=='hidden':x.set('show','0')
 elif name=='forced-wrap':sp.find('p:txBody/a:bodyPr',ns).set('wrap','none')
 elif name=='forced-align':sp.find('p:txBody/a:p/a:pPr',ns).set('algn','r')
 elif name=='forced-bold':sp.find('p:txBody/a:p/a:r/a:rPr',ns).set('b','0')
 else:E.SubElement(sp.find('p:txBody/a:p/a:r/a:rPr',ns),'{'+A+'}latin',{'typeface':'Arial'})
 d[part]=E.tostring(x)
 try:check(d)
 except (AssertionError,KeyError,AttributeError,StopIteration):negative[name]=True
 else:raise AssertionError('Accepted '+name)
src=[Path(__file__),Path('scripts/prepare-style-reset-native.mjs'),Path('scripts/research/prepare-mixed-layout-edit.py'),Path('tests/fixtures/layout-rebind.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
paths=files+['input.pdf','round3.pdf','default-materialized.pptx','default-materialized.pdf','mixed-layout-edit.pptx','mixed-layout-edit.pdf','warnings.json']
report={'scope':'current mixed run/paragraph/frame inheritance, PowerPoint 16.104, three save/close/reopen rounds','directory':str(root),'files':{p:sha(root/p) for p in paths},'sourceSha256':{str(p):sha(p) for p in src},'firstFinalPixelTextDrawingImageEquality':True,'versusMaterializedPixelEqual':[False,False,True],'bodyVerticalShiftPoints':shifts,'negativeMutationsRejected':negative,'packageOnlyLayoutEditSelectedResponded':True,'layoutEditLocalControlShiftPoints':0.96,'sharedFontMetricsShiftResolved':False,'nativeInteractiveLayoutEditVerified':False,'otherHostsVerified':False};Path('docs/testing/native-mixed-frame-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
