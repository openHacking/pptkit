#!/usr/bin/env python3
"""Independent checks for one actual PowerPoint text edit and two source saves."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
base='Edited preserved text 123 <&>';control='Unchanged control text';ui=base+' POWERPOINT UI 789';post='PPTKit after native 456 <&> POWERPOINT UI 789';note='Unchanged speaker notes';part='ppt/slides/slide1.xml';notes='ppt/notesSlides/notesSlide1.xml'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unpack(p):
 with zipfile.ZipFile(p) as z:
  assert len(z.namelist())==len(set(z.namelist()));return {n:z.read(n) for n in z.namelist()}
def shapes(data):
 x=E.fromstring(data[part]);rows={}
 for sp in x.findall('.//p:sp',ns):
  c=sp.find('p:nvSpPr/p:cNvPr',ns)
  if c is not None and c.get('id') in ('2','3'):
   assert c.get('id') not in rows;rows[c.get('id')]=sp
 assert set(rows)=={'2','3'};return rows
def shape_text(sp):return ''.join(t.text or '' for t in sp.findall('.//a:t',ns))
def check(data,expected,orphan):
 rows=shapes(data);assert shape_text(rows['2'])==expected and shape_text(rows['3'])==control
 for id,y in [('2','381000'),('3','1651000')]:
  sp=rows[id];c=sp.find('p:nvSpPr/p:cNvPr',ns);assert c.get('name')=='element-'+str(int(id)-1);x=sp.find('p:spPr/a:xfrm',ns);assert x.find('a:off',ns).attrib=={'x':'381000','y':y};assert x.find('a:ext',ns).attrib=={'cx':'8890000','cy':'889000'}
 assert [t.text for t in E.fromstring(data[notes]).iter('{'+A+'}t')]==[note];assert b'opaque retained source' in data[part];assert ('orphan.bin' in data)==orphan
files=['source.pptx','edited.pptx','interactive.pptx','round1.pptx','round2.pptx','round3.pptx','post-native-pptkit.pptx','post-round1.pptx','post-round2.pptx','post-round3.pptx'];data={n:unpack(root/n) for n in files}
check(data['source.pptx'],'Source text before',True);check(data['edited.pptx'],base,True)
for n in ('interactive.pptx','round1.pptx','round2.pptx','round3.pptx'):check(data[n],ui,False)
for n in ('post-native-pptkit.pptx','post-round1.pptx','post-round2.pptx','post-round3.pptx'):check(data[n],post,False)
def one_text_change(before,after,old,new):
 assert before.keys()==after.keys();assert {n for n in before if before[n]!=after[n]}=={part}
 x,y=E.fromstring(before[part]),E.fromstring(after[part]);target=[t for t in x.iter('{'+A+'}t') if t.text==old];assert len(target)==1;target[0].text=new;assert E.tostring(x)==E.tostring(y)
one_text_change(data['source.pptx'],data['edited.pptx'],'Source text before',base)
one_text_change(data['round3.pptx'],data['post-native-pptkit.pptx'],base,'PPTKit after native 456 <&>')
report=json.loads((root/'post-native-report.json').read_text());assert report['mode']=='edit' and report['assurance']=='scoped-preservation' and not report['losses'] and len(report['changes'])==1;assert report['changes'][0]['text']=={'before':base,'after':'PPTKit after native 456 <&>'}
# PDF text trace grouping may change when PowerPoint saves. Compare each glyph's
# identity/style and report coordinate deviations instead of asserting false equality.
pdfs={n:fitz.open(root/(n+'.pdf')) for n in ('edited','interactive','round3','post-native-pptkit','post-round3')};assert all(len(d)==1 for d in pdfs.values())
def text(p):return p.get_text()
def records(p):return p.get_pixmap(alpha=False).samples,p.get_texttrace(),p.get_drawings(),p.get_image_info(hashes=True)
def flat(p):return [(t['font'],t['size'],t['color'],c) for t in p.get_texttrace() for c in t['chars']]
for name,expected in [('edited',base),('interactive',ui),('round3',ui),('post-native-pptkit',post),('post-round3',post)]:assert expected in text(pdfs[name][0]) and control in text(pdfs[name][0])
a,b=pdfs['interactive'][0],pdfs['round3'][0];assert a.get_pixmap().samples==b.get_pixmap().samples and a.get_drawings()==b.get_drawings() and a.get_image_info(hashes=True)==b.get_image_info(hashes=True);x,y=flat(a),flat(b);assert len(x)==len(y) and all(p[:3]==q[:3] and p[3][:2]==q[3][:2] for p,q in zip(x,y));delta=max(abs(k-l) for p,q in zip(x,y) for k,l in zip(p[3][2],q[3][2]));assert 0.05<delta<0.06 and a.get_texttrace()!=b.get_texttrace()
assert records(pdfs['post-native-pptkit'][0])==records(pdfs['post-round3'][0])
def control_trace(p):
 found=[t for t in p.get_texttrace() if ''.join(chr(c[0]) for c in t['chars'])==control];assert len(found)==1;t=found[0];return (t['font'],t['size'],t['color'],[(c[0],c[2],c[3]) for c in t['chars']])
assert all(control_trace(d[0])==control_trace(pdfs['edited'][0]) for d in pdfs.values())
# Mutated candidate data must fail the independent structural checks.
negative={}
for key in ('edited-text','control-text','geometry','notes','opaque','unexpected-part'):
 d=dict(data['post-native-pptkit.pptx']);x=E.fromstring(d[part]);rows=shapes(d)
 if key in ('edited-text','control-text','geometry'):
  x=E.fromstring(d[part]);sp=next(s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('id')==('3' if key=='control-text' else '2'))
  if key=='geometry':sp.find('p:spPr/a:xfrm/a:off',ns).set('x','0')
  else:sp.find('.//a:t',ns).text='corrupted'
  d[part]=E.tostring(x)
 elif key=='notes':d[notes]=d[notes].replace(note.encode(),b'changed notes')
 elif key=='opaque':d[part]=d[part].replace(b'opaque retained source',b'other')
 else:d['unexpected.bin']=b'bad'
 try:
  check(d,post,False)
  if key=='unexpected-part':one_text_change(data['round3.pptx'],d,base,'PPTKit after native 456 <&>')
 except (AssertionError,KeyError,AttributeError,StopIteration):negative[key]=True
 else:raise AssertionError('accepted negative '+key)
manifest=json.loads(Path('docs/testing/text-edit-save-evidence.json').read_text());assert sha(root/'source.pptx')==manifest['sourceSha256'] and sha(root/'edited.pptx')==manifest['editedSha256']
paths=files+['edited.pdf','interactive.pdf','round3.pdf','post-native-pptkit.pdf','post-round3.pdf','post-native-report.json'];sources=[Path(__file__),Path('scripts/prepare-native-v1-post-edit.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/pptx-parser/src').rglob('*.ts'))
result={'scope':'one text-runs-v1 candidate, PowerPoint 16.104 interactive text edit, two actual three-round save/reopen sequences and post-native PPTKit edit','directory':str(root),'files':{p:sha(root/p) for p in paths},'sourceSha256':{str(p):sha(p) for p in sources},'historicalEvidenceSha256':sha(Path('docs/testing/text-edit-save-evidence.json')),'candidateChangedOnlySelectedText':True,'postNativeCandidateChangedOnlySelectedText':True,'notesAndControlContentGeometryRetained':True,'powerPointResaveRemovedOrphan':True,'powerPointResaveRetainedNamedOpaqueExtension':True,'interactiveAndRound3PixelsEqual':True,'interactiveAndRound3RawTextTraceEqual':False,'interactiveAndRound3MaxGlyphOriginDifferencePoints':delta,'postNativeCandidateAndRound3FullPdfRecordsEqual':True,'negativeMutationsRejected':negative,'otherHostVersionsVerified':False,'generalOpaqueReferenceSafetyVerified':False};Path('docs/testing/native-v1-interactive-evidence.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps(result,indent=2))
