#!/usr/bin/env python3
"""Independent, scoped checks of one PowerPoint V2 interactive experiment."""
import hashlib,json,posixpath,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';C='http://schemas.openxmlformats.org/drawingml/2006/chart';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';REL='http://schemas.openxmlformats.org/package/2006/relationships';ns={'p':P,'a':A,'c':C};slide='ppt/slides/slide1.xml';notes='ppt/notesSlides/notesSlide1.xml';opaque=b'opaque retained owner source';shape0='Edited preserved shape 123 <&>';note0='Edited preserved notes 456 <&>';shape1='PPTKit shape after UI <&>';note1='PPTKit notes after UI <&>'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def digest(b):return hashlib.sha256(b).hexdigest()
def unpack(p):
 with zipfile.ZipFile(p) as z:
  assert len(z.namelist())==len(set(z.namelist()));return {n:z.read(n) for n in z.namelist()}
def rels(d,part):
 dirname,name=posixpath.split(part);name=dirname+'/_rels/'+name+'.rels';x=E.fromstring(d[name]);return {c.get('Id'):c for c in x.findall('{'+REL+'}Relationship')}
def linked(d,part,typ):
 found=[c for c in rels(d,part).values() if c.get('Type','').endswith('/'+typ)];assert len(found)==1;target=posixpath.normpath(posixpath.join(posixpath.dirname(part),found[0].get('Target')));assert target in d;return target
def shape(d,part,id):
 x=E.fromstring(d[part]);found=[s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('id')==id];assert len(found)==1;return found[0]
def body_text(sp):return ''.join(n.text or '' for n in sp.findall('.//a:t',ns))
def link(d,part,sp,target,owner_text,external=False):
 refs=sp.findall('.//a:hlinkClick',ns);assert len(refs)==1;owners=[r for r in sp.findall('.//a:r',ns) if r.find('a:rPr/a:hlinkClick',ns) is not None];assert len(owners)==1 and owners[0].find('a:t',ns).text==owner_text;rid=refs[0].get('{'+R+'}id');rel=rels(d,part)[rid];assert rel.get('Target')==target and (rel.get('TargetMode')=='External')==external;assert (refs[0].get('tooltip')=='Notes link') if external else (refs[0].get('action')=='ppaction://hlinksldjump')
files=['source.pptx','edited.pptx','interactive.pptx','round1.pptx','round2.pptx','round3.pptx','post-native-pptkit.pptx','post-round1.pptx','post-round2.pptx','post-round3.pptx'];data={n:unpack(root/n) for n in files}
original=data['source.pptx'];edited=data['edited.pptx']
image_digest='45ca74602a2f7790a8eb2fad662cc73ee8ab42e78d5abfd0c66aa4d44725c9c6';workbook_digest='fefff6f4d754b43f101523daff1cb6dba01ed9bee9097d9da31099eaa2a713a9'
workbook_names=[n for n in edited if n.endswith('.xlsx')];assert len(workbook_names)==1 and digest(edited[workbook_names[0]])==workbook_digest
def check(d,shape_text,note_text,orphan):
 assert ('orphan.bin' in d)==orphan and opaque in d[slide] and opaque in d[notes]
 s=shape(d,slide,'2');assert body_text(s)==shape_text;assert body_text(shape(d,slide,'4'))=='Nested shape control';link(d,slide,s,'slide2.xml',shape_text.removesuffix(' UI SHAPE'));xfrm=s.find('p:spPr/a:xfrm',ns);assert xfrm.find('a:off',ns).attrib=={'x':'381000','y':'381000'} and xfrm.find('a:ext',ns).attrib=={'cx':'6350000','cy':'1016000'}
 x=E.fromstring(d[slide]);assert len(x.findall('.//p:pic',ns))==1 and len(x.findall('.//c:chart',ns))==1 and len(x.findall('.//a:tbl',ns))==1;assert [t.text for t in x.findall('.//a:tbl//a:t',ns)]==['Cell control A','Cell control B']
 x2=E.fromstring(d['ppt/slides/slide2.xml']);assert len(x2.findall('.//p:pic',ns))==1 and body_text(shape(d,'ppt/slides/slide2.xml','2'))=='Second slide control'
 for i in (1,2):
  sp=f'ppt/slides/slide{i}.xml';np=f'ppt/notesSlides/notesSlide{i}.xml';assert linked(d,sp,'notesSlide')==np and linked(d,np,'slide')==sp
 assert body_text(shape(d,notes,'2'))==note_text+'Notes control é🙂';assert body_text(shape(d,'ppt/notesSlides/notesSlide2.xml','2'))=='Second notes control';assert all(shape(d,f'ppt/notesSlides/notesSlide{i}.xml','2').find('p:nvSpPr/p:nvPr/p:ph',ns).attrib=={'type':'body','idx':'1'} for i in (1,2));link(d,notes,shape(d,notes,'2'),'https://example.invalid/notes?x=1&y=2',note_text.removesuffix(' UI NOTES'),True)
 image1=linked(d,slide,'image');image2=linked(d,'ppt/slides/slide2.xml','image');assert image1==image2 and digest(d[image1])==image_digest
 chart=linked(d,slide,'chart');chartroot=E.fromstring(d[chart]);values=[n.text for n in chartroot.iter('{'+C+'}v')];assert values==['Unchanged data','A','B','3','5'];workbook=linked(d,chart,'package');assert workbook.endswith('.xlsx') and digest(d[workbook])==workbook_digest
for n in ('source.pptx',):check(data[n],'Original shape text','Original speaker notes',True)
check(edited,shape0,note0,True)
for n in ('interactive.pptx','round1.pptx','round2.pptx','round3.pptx'):check(data[n],shape0+' UI SHAPE',note0+' UI NOTES',False)
for n in ('post-native-pptkit.pptx','post-round1.pptx','post-round2.pptx','post-round3.pptx'):check(data[n],shape1+' UI SHAPE',note1+' UI NOTES',False)
def two_changes(a,b,old1,new1,old2,new2):
 assert a.keys()==b.keys();assert {n for n in a if a[n]!=b[n]}=={slide,notes}
 for part,old,new in ((slide,old1,new1),(notes,old2,new2)):
  x,y=E.fromstring(a[part]),E.fromstring(b[part]);matches=[t for t in x.iter('{'+A+'}t') if t.text==old];assert len(matches)==1;matches[0].text=new;assert E.tostring(x)==E.tostring(y)
two_changes(original,edited,'Original shape text',shape0,'Original speaker notes',note0)
two_changes(data['round3.pptx'],data['post-native-pptkit.pptx'],shape0,shape1,note0,note1)
report=json.loads((root/'post-native-report.json').read_text());assert report['mode']=='edit' and report['assurance']=='scoped-preservation' and not report['losses'] and len(report['changes'])==2;assert {c['source']['part'] for c in report['changes']}=={'/'+slide,'/'+notes}
def records(p):return p.get_pixmap(alpha=False).samples,p.get_texttrace(),p.get_drawings(),p.get_image_info(hashes=True)
def flatten(p):return [(t['font'],t['size'],t['color'],ch) for t in p.get_texttrace() for ch in t['chars']]
def glyph_delta(a,b):
 x,y=flatten(a),flatten(b);assert len(x)==len(y) and all(p[:3]==q[:3] and p[3][:2]==q[3][:2] for p,q in zip(x,y));return max((abs(k-l) for p,q in zip(x,y) for k,l in zip(p[3][2],q[3][2])),default=0)
pdfs={n:fitz.open(root/(n+'.pdf')) for n in ('edited','edited-notes','interactive','interactive-notes','round3','round3-notes','post-native-pptkit','post-native-pptkit-notes','post-round3','post-round3-notes')}
assert all(len(pdfs[n])==2 for n in ('edited','interactive','round3','post-native-pptkit','post-round3'));assert all(len(pdfs[n])==3 for n in ('edited-notes','interactive-notes','round3-notes','post-native-pptkit-notes','post-round3-notes'))
assert ['Edited preserved notes 456 <&>','Notes control é🙂','Second notes control']==[p.get_text().strip() for p in pdfs['edited-notes']];assert [note0+' UI NOTES','Notes control é🙂','Second notes control']==[p.get_text().strip() for p in pdfs['round3-notes']];assert [note1+' UI NOTES','Notes control é🙂','Second notes control']==[p.get_text().strip() for p in pdfs['post-round3-notes']]
for kind in ('','-notes'):
 a,b=pdfs['post-native-pptkit'+kind],pdfs['post-round3'+kind];assert all(records(x)==records(y) for x,y in zip(a,b))
a,b=pdfs['interactive'],pdfs['round3'];assert all(x.get_pixmap().samples==y.get_pixmap().samples and x.get_image_info(hashes=True)==y.get_image_info(hashes=True) for x,y in zip(a,b));assert a[0].get_texttrace()!=b[0].get_texttrace() and a[0].get_drawings()!=b[0].get_drawings();slide_deltas=[glyph_delta(x,y) for x,y in zip(a,b)];assert 0<slide_deltas[0]<0.01 and slide_deltas[1]==0
c,d=pdfs['interactive-notes'],pdfs['round3-notes'];assert c[0].get_pixmap().samples!=d[0].get_pixmap().samples and c[0].get_image_info(hashes=True)==d[0].get_image_info(hashes=True) and all(records(c[i])==records(d[i]) for i in (1,2));notes_delta=glyph_delta(c[0],d[0]);assert 0.12<notes_delta<0.13
# Independent destructive counterexamples; no on-disk source is changed.
negative={}
for key in ('shape-text','notes-text','other-notes','chart','workbook','image','shape-link','link-owner','notes-owner','shape-geometry','opaque','orphan'):
 d=dict(data['post-native-pptkit.pptx'])
 if key in ('shape-text','notes-text','other-notes','chart','shape-geometry'):
  part={'shape-text':slide,'notes-text':notes,'other-notes':'ppt/notesSlides/notesSlide2.xml','chart':'ppt/charts/chart1.xml','shape-geometry':slide}[key];x=E.fromstring(d[part]);tag='{'+C+'}v' if key=='chart' else '{'+A+'}t'
  if key=='shape-geometry':next(s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('id')=='2').find('p:spPr/a:xfrm/a:off',ns).set('x','0')
  else:next(x.iter(tag)).text='corrupt'
  d[part]=E.tostring(x)
 elif key in ('workbook','image'):
  name=next(n for n in d if n.endswith('.xlsx')) if key=='workbook' else linked(d,slide,'image');d[name]=d[name]+b'corrupt'
 elif key=='notes-owner':
  x=E.fromstring(d[notes]);sp=next(s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('id')=='2');sp.find('p:nvSpPr/p:nvPr/p:ph',ns).set('type','sldImg');d[notes]=E.tostring(x)
 elif key=='link-owner':
  x=E.fromstring(d[slide]);sp=next(s for s in x.findall('.//p:sp',ns) if (c:=s.find('p:nvSpPr/p:cNvPr',ns)) is not None and c.get('id')=='2');runs=sp.findall('.//a:r',ns);first=runs[0].find('a:rPr',ns);second=runs[1].find('a:rPr',ns);h=first.find('a:hlinkClick',ns);first.remove(h);second.append(h);d[slide]=E.tostring(x)
 elif key=='shape-link':
  part='ppt/slides/_rels/slide1.xml.rels';x=E.fromstring(d[part]);next(c for c in x if c.get('Type','').endswith('/slide')).set('Target','missing.xml');d[part]=E.tostring(x)
 elif key=='opaque':d[notes]=d[notes].replace(opaque,b'corrupt')
 else:d['orphan.bin']=b'reappeared'
 try:check(d,shape1+' UI SHAPE',note1+' UI NOTES',False)
 except (AssertionError,KeyError,AttributeError,StopIteration):negative[key]=True
 else:raise AssertionError('accepted negative '+key)
manifest=json.loads(Path('docs/testing/text-owner-save-evidence.json').read_text());assert sha(root/'source.pptx')==manifest['sourceSha256'] and sha(root/'edited.pptx')==manifest['editedSha256']
paths=files+['edited.pdf','edited-notes.pdf','interactive.pdf','interactive-notes.pdf','round3.pdf','round3-notes.pdf','post-native-pptkit.pdf','post-native-pptkit-notes.pdf','post-round3.pdf','post-round3-notes.pdf','post-native-report.json'];sources=[Path(__file__),Path('scripts/prepare-native-v2-post-edit.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/pptx-parser/src').rglob('*.ts'))
result={'scope':'one text-runs-v2 shape/notes candidate, PowerPoint 16.104 UI edits and two three-round save/reopen sequences','directory':str(root),'files':{n:sha(root/n) for n in paths},'sourceSha256':{str(p):sha(p) for p in sources},'historicalEvidenceSha256':sha(Path('docs/testing/text-owner-save-evidence.json')),'candidateChangedOnlyTwoSelectedTextAtoms':True,'postNativeCandidateChangedOnlyTwoSelectedTextAtoms':True,'notesSlideRelationshipsAndControlsVerified':True,'linksChartWorkbookAndSharedImageVerified':True,'powerPointResaveRemovedOrphan':True,'powerPointResaveRetainedNamedOpaqueExtensions':True,'interactiveSlidePixelsEqualAfterThreeRounds':True,'interactiveSlideFirstPageRawTextAndDrawingsEqual':False,'interactiveSlideMaxGlyphOriginDeltaPoints':slide_deltas,'notesPrintPageCountBeforeAndAfter':3,'interactiveNotesFirstPagePixelsEqualAfterThreeRounds':False,'interactiveNotesFirstPageMaxGlyphOriginDeltaPoints':notes_delta,'postNativeCandidateAndRound3SlideAndNotesFullPdfRecordsEqual':True,'negativeMutationsRejected':negative,'otherHostVersionsVerified':False,'generalOpaqueReferenceSafetyVerified':False};Path('docs/testing/native-v2-interactive-evidence.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps(result,indent=2))
