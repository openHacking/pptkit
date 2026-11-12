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

probe=Path('.tmp/native-autofit-probe/run-Fxqa9h').resolve()
files=['baseline.pptx','inherit.pptx','round1.pptx','round2.pptx','round3.pptx']
data={f:unpack(root/f) for f in files}
check(data['baseline.pptx'],'baseline')
for name in files[1:]:check(data[name])
assert sha(root/'baseline.pptx')==sha(probe/'baseline.pptx')
manual=unpack(probe/'inherit.pptx');production=data['inherit.pptx'];assert manual.keys()==production.keys()
for part in manual:
 a,b=manual[part],production[part]
 if a!=b:
  try:assert E.tostring(E.fromstring(a))==E.tostring(E.fromstring(b)),part
  except E.ParseError:raise AssertionError(part)
# Mutations must be rejected independently of the positive package checks.
for page in (1,3,5):
 corrupted=dict(data['round3.pptx']);part=f'ppt/slides/slide{page}.xml';x=E.fromstring(corrupted[part]);body=x.find('.//p:sp/p:txBody/a:bodyPr',ns);E.SubElement(body,f'{{{A}}}noAutofit');corrupted[part]=E.tostring(x)
 try:check(corrupted)
 except AssertionError:pass
 else:raise AssertionError(f'accepted local override on slide {page}')
a,b=fitz.open(probe/'baseline.pdf'),fitz.open(root/'round3.pdf');assert len(a)==len(b)==6
for x,y in zip(a,b):
 assert x.get_pixmap(alpha=False).samples==y.get_pixmap(alpha=False).samples
 assert x.get_texttrace()==y.get_texttrace()
 assert x.get_drawings()==y.get_drawings()
 assert x.get_image_info(hashes=True)==y.get_image_info(hashes=True)
warning=json.loads((root/'warnings.json').read_text());assert not any(w['code']=='native-text-inheritance-materialized' for w in warning['inherit']);assert sum(w['code']=='native-text-inheritance-materialized' for w in warning['baseline'])==3
sources=[Path(__file__),Path('scripts/prepare-native-autofit-production.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'scope':'production autofit inheritance, PowerPoint 16.104, six slides, three save/close/reopen rounds','directory':str(root),'files':{f:sha(root/f) for f in files+['round3.pdf','warnings.json']},'sourceSha256':{str(p):sha(p) for p in sources},'baselinePdfSha256':sha(probe/'baseline.pdf'),'manualProbeInheritSha256':sha(probe/'inherit.pptx'),'productionPackageSemanticallyEquivalentToManualProbe':True,'baselineAndFinalPixelsTextDrawingsImagesEqual':True,'threeNativeSaveReopenRoundsVerified':True,'negativeSelectedLocalOverridesRejected':[1,3,5],'interactiveDynamicResizeProof':'manual probe only, package semantically equivalent at initial export','nativeInteractiveLayoutEditVerified':False,'otherHostsVerified':False}
Path('docs/testing/native-autofit-production-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
