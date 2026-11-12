#!/usr/bin/env python3
"""Independent individual-property controls, bounded to initial PowerPoint render."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
names=['wrap','verticalAlign','left','right','top','bottom'];attrs={'wrap':('wrap','none','square'),'verticalAlign':('anchor','t','b'),'left':('lIns','0','635000'),'right':('rIns','0','635000'),'top':('tIns','0','635000'),'bottom':('bIns','0','635000')}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unpack(p):
 with zipfile.ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
def shape(blob,name):
 x=E.fromstring(blob);items=[s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('name')==name];assert len(items)==1;return items[0]
def layouts(data):
 result={}
 for part,blob in data.items():
  if not part.startswith('ppt/slideLayouts/slideLayout') or not part.endswith('.xml'):continue
  x=E.fromstring(blob);name=x.find('p:cSld',ns).get('name');key=name.removeprefix('Target ')
  if key in attrs:assert key not in result;result[key]=part
 assert set(result)==set(names);return result
def check(data,edited=False):
 ls=layouts(data)
 for index,key in enumerate(names):
  attr,start,end=attrs[key];body=shape(data[ls[key]],'body').find('p:txBody/a:bodyPr',ns);assert body.get(attr)==(end if edited else start)
  for control in (False,True):
   page=index*2+1+control;blob=data[f'ppt/slides/slide{page}.xml'];xml=E.fromstring(blob);assert (xml.get('show')=='0')==(page==2);item=shape(blob,f'shape-{key}-'+('control' if control else 'selected'));props=item.find('p:txBody/a:bodyPr',ns)
   assert props.get(attr)==start if control else attr not in props.attrib
   assert all(other in props.attrib for other in ('wrap','anchor','lIns','rIns','tIns','bIns') if control or other!=attr)
   assert props.find('a:noAutofit',ns) is not None
   assert item.find('p:txBody/a:p/a:r/a:t',ns).text
source=unpack(root/'inherit.pptx');base=unpack(root/'materialize.pptx');edited=unpack(root/'layout-edit.pptx');rebuilt=unpack(root/'rebuilt-edit.pptx');check(source);check(edited,True);assert edited==rebuilt
for name in ('round1.pptx','round2.pptx','round3.pptx'):check(unpack(root/name))
ls=layouts(source);assert source.keys()==edited.keys();assert all(source[p]==edited[p] for p in source if p not in ls.values())
for key in names:
 original=E.fromstring(source[ls[key]]);changed=E.fromstring(edited[ls[key]]);attr,start,end=attrs[key]
 # Compare the full tree after undoing the only authorized attribute change.
 for s in changed.findall('.//p:sp',ns):
  c=s.find('p:nvSpPr/p:cNvPr',ns)
  if c is not None and c.get('name')=='body':s.find('p:txBody/a:bodyPr',ns).set(attr,start)
 assert E.tostring(original)==E.tostring(changed)
for i,key in enumerate(names):
 attr,start,_=attrs[key];p=shape(base[f'ppt/slides/slide{i*2+1}.xml'],f'shape-{key}-selected').find('p:txBody/a:bodyPr',ns);assert p.get(attr)==start
pdfs=[fitz.open(root/(n+'.pdf')) for n in ('materialize','inherit','layout-edit','round3')];assert all(len(d)==12 for d in pdfs)
def normalized_trace(page):
 # Font subset glyph numbers are PDF-local; compare Unicode, positions and all other fields.
 return [{**t,'chars':[(c[0],*c[2:]) for c in t['chars']]} for t in page.get_texttrace()]
raw_gid_differences=[]
labels=['Wrapping changes','Vertical alignment target','Left margin target','Right margin target','Top margin target','Bottom margin target']
for i in range(12):
 a,b,c,final=(d[i] for d in pdfs);assert a.get_pixmap(alpha=False).samples==b.get_pixmap(alpha=False).samples==final.get_pixmap(alpha=False).samples;assert a.get_texttrace()==b.get_texttrace()==final.get_texttrace();assert a.get_drawings()==b.get_drawings()==final.get_drawings();assert a.get_image_info(hashes=True)==b.get_image_info(hashes=True)==final.get_image_info(hashes=True)
 assert all(labels[i//2] in page.get_text() for page in (a,b,c,final))
 equal=b.get_pixmap(alpha=False).samples==c.get_pixmap(alpha=False).samples;assert equal==(i%2==1)
 if i%2:
  assert normalized_trace(b)==normalized_trace(c)
  if b.get_texttrace()!=c.get_texttrace():raw_gid_differences.append(i+1)
 else:assert normalized_trace(b)!=normalized_trace(c)
negative=root/'negative';negative.mkdir(exist_ok=True);results={}
for key in names:
 data=dict(source);part=ls[key];x=E.fromstring(data[part]);attr=attrs[key][0]
 for s in x.findall('.//p:sp',ns):
  c=s.find('p:nvSpPr/p:cNvPr',ns)
  if c is not None and c.get('name')=='body':s.find('p:txBody/a:bodyPr',ns).set(attr,'invalid')
 data[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);file=negative/(key+'.pptx')
 with zipfile.ZipFile(file,'w',zipfile.ZIP_DEFLATED) as z:
  for n,v in data.items():z.writestr(n,v)
 try:check(unpack(file));raise RuntimeError('negative accepted: '+key)
 except AssertionError:results[key]=sha(file)
files=['materialize.pptx','inherit.pptx','round1.pptx','round2.pptx','round3.pptx','layout-edit.pptx','rebuilt-edit.pptx','materialize.pdf','inherit.pdf','round3.pdf','layout-edit.pdf','cases.json'];sources=[Path(__file__),Path('scripts/prepare-native-frame-individual.mjs'),Path('scripts/research/prepare-native-frame-individual-edit.py'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'scope':'six separately selected frame attributes, three PowerPoint 16.104 save/reopen rounds, 12-page initial/final PDFs, package-level layout changes','directory':str(root),'files':{f:sha(root/f) for f in files},'sourceSha256':{str(p):sha(p) for p in sources},'baselineMatchesAllPages':True,'layoutChangesSelectedPagesOnly':True,'controlNormalizedTextTraceMatches':True,'rawGlyphSubsetIdDifferencesOnControlPages':raw_gid_differences,'negativeZip':results,'saveReopenVerified':True,'interactiveLayoutEditVerified':False,'otherHostsVerified':False};Path('docs/testing/native-frame-individual-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
