#!/usr/bin/env python3
"""Independent checks of a named PowerPoint merged-table editing control."""
import hashlib,json,posixpath,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';C='http://schemas.openxmlformats.org/drawingml/2006/chart';REL='http://schemas.openxmlformats.org/package/2006/relationships';ns={'p':P,'a':A,'c':C};slide='ppt/slides/slide1.xml';image_hash='45ca74602a2f7790a8eb2fad662cc73ee8ab42e78d5abfd0c66aa4d44725c9c6';workbook_hash='fefff6f4d754b43f101523daff1cb6dba01ed9bee9097d9da31099eaa2a713a9'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def digest(b):return hashlib.sha256(b).hexdigest()
def unpack(p):
 with zipfile.ZipFile(p) as z:
  assert len(z.namelist())==len(set(z.namelist()));return {n:z.read(n) for n in z.namelist()}
def linked(d,part,typ):
 dirname,name=posixpath.split(part);x=E.fromstring(d[dirname+'/_rels/'+name+'.rels']);found=[c for c in x.findall('{'+REL+'}Relationship') if c.get('Type','').endswith('/'+typ)];assert len(found)==1;target=posixpath.normpath(posixpath.join(dirname,found[0].get('Target')));assert target in d;return target
def table(d,id):
 x=E.fromstring(d[slide]);frames=[f for f in x.findall('.//p:graphicFrame',ns) if (c:=f.find('p:nvGraphicFramePr/p:cNvPr',ns)) is not None and c.get('id')==id];assert len(frames)==1;t=frames[0].find('.//a:tbl',ns);assert t is not None;return t
def cells(t):return [r.findall('a:tc',ns) for r in t.findall('a:tr',ns)]
def txt(x):return ''.join(n.text or '' for n in x.findall('.//a:t',ns))
def check(d,anchor,right,native):
 t=table(d,'8');assert [c.get('w') for c in t.findall('a:tblGrid/a:gridCol',ns)]==['889000','889000','1143000'];assert [r.get('h') for r in t.findall('a:tr',ns)]==['381000']*3;c=cells(t);assert [len(r) for r in c]==[3,3,3]
 expected=[[{'gridSpan':'2','rowSpan':'2'},{'hMerge':'1'},{}],[{'vMerge':'1'},{'vMerge':'1','hMerge':'1'},{}],[{},{},{}]] if not native else [[{'gridSpan':'2','rowSpan':'2'},{'hMerge':'1','rowSpan':'2'},{}],[{'vMerge':'1','gridSpan':'2'},{'vMerge':'1','hMerge':'1'},{}],[{},{},{}]]
 assert [[x.attrib for x in row] for row in c]==expected;assert [[txt(x) for x in row] for row in c]==[[anchor,'','Top right control'],['','',right],['Last row source','','Last cell control']];assert [[txt(x) for x in row] for row in cells(table(d,'6'))]==[['Cell control A','Cell control B']]
 assert ('orphan.bin' in d)==(not native) and b'opaque table source' in d[slide];unknown=[v for tc in E.fromstring(d[slide]).findall('.//a:tcPr',ns) for k,v in tc.attrib.items() if k=='{urn:future}opaque'];assert ('cell-source' in unknown)==(not native)
 for i,expected_notes in ((1,['Original speaker notes','Notes control é🙂']),(2,['Second notes control'])):
  sp=f'ppt/slides/slide{i}.xml';np=f'ppt/notesSlides/notesSlide{i}.xml';assert linked(d,sp,'notesSlide')==np and linked(d,np,'slide')==sp;assert [x.text for x in E.fromstring(d[np]).iter('{'+A+'}t')]==expected_notes
 image1=linked(d,slide,'image');image2=linked(d,'ppt/slides/slide2.xml','image');assert image1==image2 and digest(d[image1])==image_hash
 chart=linked(d,slide,'chart');assert [v.text for v in E.fromstring(d[chart]).iter('{'+C+'}v')]==['Unchanged data','A','B','3','5'];book=linked(d,chart,'package');assert book.endswith('.xlsx') and digest(d[book])==workbook_hash
files=['source.pptx','edited.pptx','interactive.pptx','round1.pptx','round2.pptx','round3.pptx','post-native-pptkit.pptx','post-round1.pptx','post-round2.pptx','post-round3.pptx'];data={n:unpack(root/n) for n in files}
check(data['source.pptx'],'Merged anchor control','Bottom right source',False);check(data['edited.pptx'],'Anchor 123','Cell 456',False)
for n in ('interactive.pptx','round1.pptx','round2.pptx','round3.pptx'):check(data[n],'Anchor 123 UIA','Cell 456 UIB',True)
for n in ('post-native-pptkit.pptx','post-round1.pptx','post-round2.pptx','post-round3.pptx'):check(data[n],'Anchor PPTKit UIA','Cell PPTKit UIB',True)
def changes(a,b,pairs):
 assert a.keys()==b.keys() and {n for n in a if a[n]!=b[n]}=={slide};x,y=E.fromstring(a[slide]),E.fromstring(b[slide])
 for old,new in pairs:
  nodes=[n for n in x.iter('{'+A+'}t') if n.text==old];assert len(nodes)==1;nodes[0].text=new
 assert E.tostring(x)==E.tostring(y)
changes(data['source.pptx'],data['edited.pptx'],[('Merged anchor control','Anchor 123'),('Bottom right source','Cell 456')]);changes(data['round3.pptx'],data['post-native-pptkit.pptx'],[('Anchor 123 UIA','Anchor PPTKit UIA'),('Cell 456 UIB','Cell PPTKit UIB')])
report=json.loads((root/'post-native-report.json').read_text());assert report['mode']=='edit' and report['assurance']=='scoped-preservation' and not report['losses'] and len(report['changes'])==2;assert {c['text']['after'] for c in report['changes']}=={'Anchor PPTKit UIA','Cell PPTKit UIB'};assert {'/'.join(c['source']['path'].split('/')[-2:]) for c in report['changes']}=={'a:tr[1]/a:tc[1]','a:tr[2]/a:tc[3]'} and all(c['source']['sourceObjectId']=='8' for c in report['changes']);assert {(c['text']['before'],c['text']['after']) for c in report['changes']}=={('Anchor 123 UIA','Anchor PPTKit UIA'),('Cell 456 UIB','Cell PPTKit UIB')}
def rec(p):return p.get_pixmap(alpha=False).samples,p.get_texttrace(),p.get_drawings(),p.get_image_info(hashes=True)
def glyphs(p):return [(t['font'],t['size'],t['color'],c) for t in p.get_texttrace() for c in t['chars']]
pdfs={n:fitz.open(root/(n+'.pdf')) for n in ('edited','interactive','round3','post-native-pptkit','post-final')};assert all(len(d)==2 for d in pdfs.values())
for n,words in [('edited',('Anchor123','Cell456')),('interactive',('Anchor123UIA','Cell456UIB')),('round3',('Anchor123UIA','Cell456UIB')),('post-native-pptkit',('AnchorPPTKitUIA','CellPPTKitUIB')),('post-final',('AnchorPPTKitUIA','CellPPTKitUIB'))]:
 doc=pdfs[n];text=''.join(doc[0].get_text().split());assert all(w in text for w in words) and 'CellcontrolA' in text and 'CellcontrolB' in text;assert 'Second slide control' in doc[1].get_text()
a,b=pdfs['interactive'],pdfs['round3'];assert all(x.get_pixmap().samples==y.get_pixmap().samples and x.get_drawings()==y.get_drawings() and x.get_image_info(hashes=True)==y.get_image_info(hashes=True) for x,y in zip(a,b));assert a[0].get_texttrace()!=b[0].get_texttrace() and rec(a[1])==rec(b[1]);x,y=glyphs(a[0]),glyphs(b[0]);assert len(x)==len(y) and all(p[:3]==q[:3] and p[3][:2]==q[3][:2] for p,q in zip(x,y));delta=max(abs(k-l) for p,q in zip(x,y) for k,l in zip(p[3][2],q[3][2]));assert 0<delta<0.02
assert all(rec(x)==rec(y) for x,y in zip(pdfs['post-native-pptkit'],pdfs['post-final']))
negative={}
for key in ('anchor','right','grid','merge','control','notes','image','chart','workbook','opaque','cell-attribute','orphan','extra-part'):
 d=dict(data['post-native-pptkit.pptx'])
 if key in ('anchor','right','grid','merge','control'):
  x=E.fromstring(d[slide]);f=next(f for f in x.findall('.//p:graphicFrame',ns) if (c:=f.find('p:nvGraphicFramePr/p:cNvPr',ns)) is not None and c.get('id')=='8');t=f.find('.//a:tbl',ns);c=cells(t)
  if key=='anchor':c[0][0].find('.//a:t',ns).text='wrong'
  elif key=='right':c[1][2].find('.//a:t',ns).text='wrong'
  elif key=='grid':t.find('a:tblGrid/a:gridCol',ns).set('w','1')
  elif key=='merge':c[1][1].set('hMerge','0')
  else:c[0][2].find('.//a:t',ns).text='wrong'
  d[slide]=E.tostring(x)
 elif key=='notes':d['ppt/notesSlides/notesSlide2.xml']=d['ppt/notesSlides/notesSlide2.xml'].replace(b'Second notes control',b'wrong')
 elif key in ('image','workbook'):
  name=linked(d,slide,'image') if key=='image' else linked(d,linked(d,slide,'chart'),'package');d[name]+=b'wrong'
 elif key=='chart':
  name=linked(d,slide,'chart');x=E.fromstring(d[name]);next(x.iter('{'+C+'}v')).text='wrong';d[name]=E.tostring(x)
 elif key=='opaque':d[slide]=d[slide].replace(b'opaque table source',b'wrong')
 elif key=='cell-attribute':
  x=E.fromstring(d[slide]);next(x.iter('{'+A+'}tcPr')).set('{urn:future}opaque','cell-source');d[slide]=E.tostring(x)
 elif key=='orphan':d['orphan.bin']=b'wrong'
 else:d['extra.bin']=b'wrong'
 try:
  check(d,'Anchor PPTKit UIA','Cell PPTKit UIB',True)
  if key=='extra-part':changes(data['round3.pptx'],d,[('Anchor 123 UIA','Anchor PPTKit UIA'),('Cell 456 UIB','Cell PPTKit UIB')])
 except (AssertionError,KeyError,AttributeError,StopIteration):negative[key]=True
 else:raise AssertionError('accepted negative '+key)
history=json.loads(Path('docs/testing/table-save-evidence.json').read_text());assert sha(root/'source.pptx')==history['sourceSha256'] and sha(root/'edited.pptx')==history['editedSha256']
paths=files+['edited.pdf','interactive.pdf','round3.pdf','post-native-pptkit.pdf','post-final.pdf','post-native-report.json'];sources=[Path(__file__),Path('scripts/prepare-native-v3-post-edit.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/pptx-parser/src').rglob('*.ts'))
result={'scope':'one text-runs-v3 merged-table candidate, PowerPoint 16.104 UI edits and two three-round save/reopen sequences','directory':str(root),'files':{n:sha(root/n) for n in paths},'sourceSha256':{str(p):sha(p) for p in sources},'historicalEvidenceSha256':sha(Path('docs/testing/table-save-evidence.json')),'candidateChangedOnlyTwoSelectedTextAtoms':True,'postNativeCandidateChangedOnlyTwoSelectedTextAtoms':True,'physicalGridAndMergeSemanticsVerified':True,'notesControlsImageChartWorkbookVerified':True,'powerPointResaveRemovedOrphanAndUnknownCellAttribute':True,'powerPointResaveRetainedNamedOpaqueExtension':True,'powerPointRewroteMergeContinuationAttributes':True,'interactiveAndRound3PixelsEqual':True,'interactiveFirstPageRawTextTraceEqual':False,'interactiveFirstPageMaxGlyphOriginDifferencePoints':delta,'postNativeCandidateAndRound3FullPdfRecordsEqual':True,'negativeMutationsRejected':negative,'otherHostVersionsVerified':False,'generalOpaqueReferenceSafetyVerified':False};Path('docs/testing/native-v3-interactive-evidence.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps(result,indent=2))
