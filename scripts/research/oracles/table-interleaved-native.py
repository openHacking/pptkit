#!/usr/bin/env python3
"""Independent OPC/XML/PDF checks of named headless interleaved table saves."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';C='http://schemas.openxmlformats.org/drawingml/2006/chart';ns={'p':P,'a':A,'c':C};slide='ppt/slides/slide1.xml';image_hash='45ca74602a2f7790a8eb2fad662cc73ee8ab42e78d5abfd0c66aa4d44725c9c6'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def digest(b):return hashlib.sha256(b).hexdigest()
def unpack(p):
 with zipfile.ZipFile(p) as z:
  assert len(z.namelist())==len(set(z.namelist()));return {n:z.read(n) for n in z.namelist()}
def table(d,id):
 x=E.fromstring(d[slide]);frames=[f for f in x.findall('.//p:graphicFrame',ns) if (c:=f.find('p:nvGraphicFramePr/p:cNvPr',ns)) is not None and c.get('id')==id];assert len(frames)==1;t=frames[0].find('.//a:tbl',ns);assert t is not None;return t
def txt(x):return ''.join(n.text or '' for n in x.findall('.//a:t',ns))
def cells(t):return [r.findall('a:tc',ns) for r in t.findall('a:tr',ns)]
def check(d,anchor,right,expected_size):
 t=table(d,'16');assert [c.get('w') for c in t.findall('a:tblGrid/a:gridCol',ns)]==['888840','888840','1143000'];assert [r.get('h') for r in t.findall('a:tr',ns)]==['380880']*3;c=cells(t);assert [len(r) for r in c]==[3,3,3]
 assert [[x.attrib for x in row] for row in c]==[[{'gridSpan':'2','rowSpan':'2'},{'hMerge':'1','rowSpan':'1'},{}],[{'vMerge':'1','gridSpan':'1'},{'vMerge':'1','hMerge':'1'},{}],[{},{},{}]]
 assert [[txt(x) for x in row] for row in c]==[[anchor,'','Top right control'],['','',right],['Last row source','','Last cell control']]
 control=table(d,'14');assert [[txt(x) for x in row] for row in cells(control)]==[['Cell control A','Cell control B']]
 for i,expected in ((1,['Original speaker notes','Notes control é🙂']),(2,['Second notes control'])):
  p=f'ppt/notesSlides/notesSlide{i}.xml';assert [x.text for x in E.fromstring(d[p]).iter('{'+A+'}t')]==expected
 s2=E.fromstring(d['ppt/slides/slide2.xml']);assert 'Second slide control' in [x.text for x in s2.iter('{'+A+'}t')];pics=s2.findall('.//p:pic',ns);assert len(pics)==1;transform=pics[0].find('p:spPr/a:xfrm',ns);assert transform.find('a:off',ns).attrib=={'x':'8255160','y':'4444920'};assert tuple(int(transform.find('a:ext',ns).get(k)) for k in ('cx','cy'))==expected_size
 images=[n for n in d if n.startswith('ppt/media/')];assert len(images)==1 and digest(d[images[0]])==image_hash
 charts=[n for n in d if n.startswith('ppt/charts/chart') and n.endswith('.xml')];assert len(charts)==1;assert [x.text for x in E.fromstring(d[charts[0]]).iter('{'+C+'}v')]==['Unchanged data','A','B','3','5'];assert not any(n.endswith('.xlsx') for n in d);assert 'orphan.bin' not in d
files=['source.pptx']+[f'{prefix}round{i}.pptx' for i in (1,2,3) for prefix in ('','native/')];data={n:unpack(root/n) for n in files};source=data['source.pptx'];history=json.loads(Path('docs/testing/table-save-evidence.json').read_text());prior=unpack(root/'prior-pptkit-candidate.pptx');assert sha(root/'source.pptx')==history['nativeSha256'] and sha(root/'prior-pptkit-candidate.pptx')==history['editedSha256'];assert 'orphan.bin' in prior and any(n.endswith('.xlsx') for n in prior) and b'opaque table source' in prior[slide];assert b'opaque table source' not in source[slide];check(source,'Anchor 123','Cell 456',(1269720,761760))
sizes=[];diffs=[]
for i in (1,2,3):
 before=source if i==1 else data[f'native/round{i-1}.pptx'];candidate=data[f'round{i}.pptx'];native=data[f'native/round{i}.pptx'];old_anchor='Anchor 123' if i==1 else f'Interleaved anchor {i-1}';old_right='Cell 456' if i==1 else f'Interleaved cell {i-1}';anchor=f'Interleaved anchor {i}';right=f'Interleaved cell {i}';before_size=(1269720-360*(i-1),761760-360*(i-1));after_size=(before_size[0]-360,before_size[1]-360)
 check(before,old_anchor,old_right,before_size);check(candidate,anchor,right,before_size);check(native,anchor,right,after_size);sizes.append({'round':i,'pptkit':before_size,'native':after_size})
 assert before.keys()==candidate.keys() and {n for n in before if before[n]!=candidate[n]}=={slide};a,b=E.fromstring(before[slide]),E.fromstring(candidate[slide]);
 for old,new in ((old_anchor,anchor),(old_right,right)):
  found=[n for n in a.iter('{'+A+'}t') if n.text==old];assert len(found)==1;found[0].text=new
 assert E.tostring(a)==E.tostring(b)
 report=json.loads((root/f'round{i}.json').read_text());assert report['mode']=='edit' and report['assurance']=='scoped-preservation' and not report['losses'] and len(report['changes'])==2;assert {x['text']['after'] for x in report['changes']}=={anchor,right}
 d1=fitz.open(root/f'pdf/round{i}.pdf');d2=fitz.open(root/f'reopened-pdf/round{i}.pdf');assert len(d1)==len(d2)==2
 for doc in (d1,d2):
  text=''.join(doc[0].get_text().split());assert f'Interleavedanchor{i}' in text and f'Interleavedcell{i}' in text and 'CellcontrolA' in text and 'CellcontrolB' in text;assert 'Second slide control' in doc[1].get_text()
 same=[x.get_pixmap(alpha=False).samples==y.get_pixmap(alpha=False).samples for x,y in zip(d1,d2)];assert same==[False,True];diffs.append({'round':i,'slidePixelsEqualCandidateVsNative':same,'imageRecordsEqual':[x.get_image_info(hashes=True)==y.get_image_info(hashes=True) for x,y in zip(d1,d2)]});assert diffs[-1]['imageRecordsEqual']==[False,False]
# Independent mutation probes demonstrate failures for wrong physical anchors,
# topology, controls, resources, notes and unintended package mutations.
negative={}
for key in ('anchor','right','grid','merge','control','notes','image','chart','image-geometry','orphan','extra-part'):
 d=dict(data['round3.pptx']);part=slide
 if key in ('anchor','right','grid','merge','control'):
  x=E.fromstring(d[part]);frame=next(f for f in x.findall('.//p:graphicFrame',ns) if (c:=f.find('p:nvGraphicFramePr/p:cNvPr',ns)) is not None and c.get('id')=='16');t=frame.find('.//a:tbl',ns);c=cells(t)
  if key=='anchor':c[0][0].find('.//a:t',ns).text='wrong'
  elif key=='right':c[1][2].find('.//a:t',ns).text='wrong'
  elif key=='grid':t.find('a:tblGrid/a:gridCol',ns).set('w','1')
  elif key=='merge':c[1][1].set('hMerge','0')
  else:c[0][2].find('.//a:t',ns).text='wrong'
  d[part]=E.tostring(x)
 elif key=='notes':d['ppt/notesSlides/notesSlide2.xml']=d['ppt/notesSlides/notesSlide2.xml'].replace(b'Second notes control',b'wrong')
 elif key=='image':name=next(n for n in d if n.startswith('ppt/media/'));d[name]+=b'wrong'
 elif key=='chart':
  name=next(n for n in d if n.startswith('ppt/charts/chart') and n.endswith('.xml'));x=E.fromstring(d[name]);next(x.iter('{'+C+'}v')).text='wrong';d[name]=E.tostring(x)
 elif key=='image-geometry':
  part='ppt/slides/slide2.xml';x=E.fromstring(d[part]);x.find('.//p:pic/p:spPr/a:xfrm/a:ext',ns).set('cx','1');d[part]=E.tostring(x)
 elif key=='orphan':d['orphan.bin']=b'wrong'
 else:d['extra.bin']=b'wrong'
 try:
  check(d,'Interleaved anchor 3','Interleaved cell 3',(1269000,761040))
  if key=='extra-part':assert d.keys()==data['native/round2.pptx'].keys()
 except (AssertionError,KeyError,AttributeError,StopIteration):negative[key]=True
 else:raise AssertionError('accepted negative '+key)
metadata=json.loads((root/'headless-run.json').read_text());reopen=json.loads((root/'reopen-run.json').read_text());assert metadata['version']==reopen['version']=='LibreOfficeDev 26.8.0.0.alpha0 2c87e51eeaa2b413ff4ae097b2705eea1995d8e5';assert len(metadata['runs'])==len(reopen['rounds'])==3
paths=files+['prior-pptkit-candidate.pptx']+[f'{p}/round{i}.pdf' for i in (1,2,3) for p in ('pdf','reopened-pdf')]+[f'round{i}.json' for i in (1,2,3)]+['headless-run.json','reopen-run.json'];sources=[Path(__file__),Path('scripts/prepare-table-interleaved-native.mjs'),Path('scripts/inspect-table-native-reopens.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/pptx-parser/src').rglob('*.ts'))
result={'scope':'three interleaved text-runs-v3 edits and LibreOfficeDev headless save/reopen cycles starting from a prior native resave','directory':str(root),'files':{n:sha(root/n) for n in paths},'sourceSha256':{str(p):sha(p) for p in sources},'historicalEvidenceSha256':sha(Path('docs/testing/table-save-evidence.json')),'pptkitEachRoundOnlyTwoSelectedAtomsChanged':True,'physicalMergeAndGridVerified':True,'sourceNativeAlreadyLostWorkbookAndOrphan':True,'sourceNativeAlreadyLostNamedOpaqueExtension':True,'nativeImageGeometryShrinkPerResaveEmu':[360,360],'imageGeometryByRound':sizes,'pdfCandidateVsNative':diffs,'negativeMutationsRejected':negative,'powerPointInteractiveVerified':False,'nativeFontAndWholeAppearanceEqualityVerified':False};Path('docs/testing/table-interleaved-native-evidence.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps(result,indent=2))
