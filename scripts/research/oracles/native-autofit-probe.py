#!/usr/bin/env python3
"""Independent PowerPoint autofit omission probe, including counterexample."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
names=('none','shrink','resize');tags={'none':('noAutofit',{}),'shrink':('normAutofit',{'fontScale':'60000'}),'resize':('spAutoFit',{})};edit_tags={'none':tags['shrink'],'shrink':tags['none'],'resize':tags['none']};fit={f'{{{A}}}{x}' for x in ('noAutofit','normAutofit','spAutoFit')}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unpack(p):
 with zipfile.ZipFile(p) as z:return {i.filename:z.read(i.filename) for i in z.infolist()}
def shape(blob,name):
 x=E.fromstring(blob);rows=[s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('name')==name];assert len(rows)==1;return rows[0]
def layouts(d):
 out={}
 for part,blob in d.items():
  if not part.startswith('ppt/slideLayouts/slideLayout') or not part.endswith('.xml'):continue
  x=E.fromstring(blob);key=x.find('p:cSld',ns).get('name').removeprefix('Target ')
  if key in names:assert key not in out;out[key]=part
 assert set(out)==set(names);return out
def child(b):
 found=[(x.tag.split('}')[-1],x.attrib) for x in b if x.tag in fit];assert len(found)<=1;return found[0] if found else None
def check(d,mode='inherit'):
 ls=layouts(d)
 for j,key in enumerate(names):
  body=shape(d[ls[key]],'body').find('p:txBody/a:bodyPr',ns);assert child(body)==(edit_tags[key] if mode=='edit' else tags[key])
  for control in (False,True):
   page=j*2+1+control;part=d[f'ppt/slides/slide{page}.xml'];slide=E.fromstring(part);assert (slide.get('show')=='0')==(page==2);sp=shape(part,f'shape-{key}-'+('control' if control else 'selected'));body=sp.find('p:txBody/a:bodyPr',ns);assert child(body)==(tags[key] if mode=='baseline' or control else None);assert key+' long text:' in ''.join(sp.itertext())
files=['baseline.pptx','inherit.pptx','round1.pptx','round2.pptx','round3.pptx','edit.pptx'];data={f:unpack(root/f) for f in files};check(data['baseline.pptx'],'baseline')
for name in ('inherit.pptx','round1.pptx','round2.pptx','round3.pptx'):check(data[name])
check(data['edit.pptx'],'edit');original=data['inherit.pptx'];edited=data['edit.pptx'];ls=layouts(original);assert original.keys()==edited.keys();assert all(original[p]==edited[p] for p in original if p not in ls.values())
# Verify that each target layout changed only the intended autofit child.
for key in names:
 orig=E.fromstring(original[ls[key]]);changed=E.fromstring(edited[ls[key]])
 for s in changed.findall('.//p:sp',ns):
  c=s.find('p:nvSpPr/p:cNvPr',ns)
  if c is not None and c.get('name')=='body':
   b=s.find('p:txBody/a:bodyPr',ns);old=[x for x in b if x.tag in fit];assert len(old)==1;b.remove(old[0]);tag,attrib=tags[key];E.SubElement(b,f'{{{A}}}{tag}',attrib)
 assert E.tostring(orig)==E.tostring(changed)
pdfs=[fitz.open(root/(n+'.pdf')) for n in ('baseline','inherit','round3','edit')];assert all(len(d)==6 for d in pdfs)
def trace(page):return [{**t,'chars':[(c[0],*c[2:]) for c in t['chars']]} for t in page.get_texttrace()]
raw_diff=[]
for i in range(6):
 a,b,c,d=[doc[i] for doc in pdfs];assert a.get_pixmap(alpha=False).samples==b.get_pixmap(alpha=False).samples==c.get_pixmap(alpha=False).samples;assert a.get_texttrace()==b.get_texttrace()==c.get_texttrace();assert a.get_drawings()==b.get_drawings()==c.get_drawings();assert a.get_image_info(hashes=True)==b.get_image_info(hashes=True)==c.get_image_info(hashes=True)
 equal=b.get_pixmap(alpha=False).samples==d.get_pixmap(alpha=False).samples;assert equal==(i not in (0,2));assert (trace(b)==trace(d))==(i not in (0,2));assert all(names[i//2]+' long text:' in p.get_text() for p in (a,b,c,d))
 if equal and b.get_texttrace()!=d.get_texttrace():raw_diff.append(i+1)
# Actual interactive text additions: inherited resize and explicit resize both grow;
# an otherwise identical copy whose layout was switched to none keeps its height.
live=unpack(root/'interactive-resize.pptx');live_none=unpack(root/'interactive-none.pptx')
def dynamic(d,page,name):
 sp=shape(d[f'ppt/slides/slide{page}.xml'],name);ext=sp.find('p:spPr/a:xfrm/a:ext',ns);body=sp.find('p:txBody/a:bodyPr',ns);return (int(ext.get('cx')),int(ext.get('cy'))),child(body),''.join(sp.itertext())
base_selected=dynamic(data['round3.pptx'],5,'shape-resize-selected');base_control=dynamic(data['round3.pptx'],6,'shape-resize-control')
selected=dynamic(live,5,'shape-resize-selected');control=dynamic(live,6,'shape-resize-control');none=dynamic(live_none,5,'shape-resize-selected');none_control=dynamic(live_none,6,'shape-resize-control')
assert base_selected[0]==base_control[0]==none[0]==none_control[0]==(2667000,889000)
assert selected[0]==control[0] and selected[0][0]==2667000 and selected[0][1]>889000
assert selected[1] is None and control[1]==tags['resize'] and none[1] is None and none_control[1]==tags['resize']
assert selected[2]==control[2]==none[2] and selected[2].startswith(base_selected[2]) and len(selected[2])>len(base_selected[2]);assert none_control[2]==base_control[2]
assert child(shape(live[layouts(live)['resize']],'body').find('p:txBody/a:bodyPr',ns))==tags['resize']
assert child(shape(live_none[layouts(live_none)['resize']],'body').find('p:txBody/a:bodyPr',ns))==tags['none']
negative=root/'negative';negative.mkdir(exist_ok=True);results={}
for key in names:
 d=dict(data['round3.pptx']);part=layouts(d)[key];x=E.fromstring(d[part]);b=x.find('.//p:sp/p:txBody/a:bodyPr',ns);old=[c for c in b if c.tag in fit];assert len(old)==1;b.remove(old[0]);E.SubElement(b,f'{{{A}}}spAutoFit' if key!='resize' else f'{{{A}}}noAutofit');d[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);file=negative/(key+'.pptx')
 with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED) as z:
  for n,v in d.items():z.writestr(n,v)
 try:check(unpack(file));raise RuntimeError('negative accepted '+key)
 except AssertionError:results[key]=sha(file)
for key in ('selected-local','control-mode','hidden'):
 d=dict(data['round3.pptx']);part='ppt/slides/slide2.xml' if key=='hidden' else 'ppt/slides/slide3.xml' if key=='selected-local' else 'ppt/slides/slide4.xml';x=E.fromstring(d[part])
 if key=='hidden':x.attrib.pop('show',None)
 else:
  b=x.find('.//p:sp/p:txBody/a:bodyPr',ns)
  for old in [c for c in b if c.tag in fit]:b.remove(old)
  E.SubElement(b,f'{{{A}}}spAutoFit')
 d[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);file=negative/(key+'.pptx')
 with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED) as z:
  for n,v in d.items():z.writestr(n,v)
 try:check(unpack(file));raise RuntimeError('negative accepted '+key)
 except AssertionError:results[key]=sha(file)
paths=files+['baseline.pdf','inherit.pdf','round3.pdf','edit.pdf','warnings.json','interactive-resize.pptx','interactive-none.pptx'];sources=[Path(__file__),Path('scripts/prepare-native-autofit-probe.mjs'),Path('scripts/research/prepare-native-autofit-omission.py'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'scope':'manual autofit omission, 3 modes, 6 slides, PowerPoint 16.104 three save/reopen rounds; package-only layout edits and actual text edits','directory':str(root),'files':{f:sha(root/f) for f in paths},'sourceSha256':{str(p):sha(p) for p in sources},'baselineInitialFinalEquality':True,'layoutEditChangedSelectedPages':[1,3],'layoutEditDidNotVisiblyChangeResizePage':5,'localControlPixelsAndNormalizedTextEqual':True,'rawGlyphSubsetIdDifferencesOnEqualPages':raw_diff,'negativeZip':results,'dynamicResizeAfterTextEditVerified':True,'dynamicResizeHeightEmu':selected[0][1],'dynamicNoneHeightEmu':none[0][1],'productionNativeAutofitExportVerified':False,'nativeInteractiveLayoutEditVerified':False,'otherHostsVerified':False};Path('docs/testing/native-autofit-probe-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
