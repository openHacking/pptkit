#!/usr/bin/env python3
"""Independent, narrowly scoped checks of the actual PowerPoint selective-reset control."""
import hashlib, json, sys, zipfile, xml.etree.ElementTree as ET
from pathlib import Path
import fitz
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'; R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'; REL='http://schemas.openxmlformats.org/package/2006/relationships'
ns={'p':P,'a':A,'r':R,'rel':REL}
root=Path(sys.argv[1]).resolve()
def digest(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def owner(entries,part,kind):
 folder,name=part.rsplit('/',1); relationships=ET.fromstring(entries[f'{folder}/_rels/{name}.rels']); edge=[x for x in relationships if x.get('Type','').endswith('/'+kind)]; assert len(edge)==1
 import posixpath
 return posixpath.normpath(posixpath.join(folder,edge[0].get('Target')))
def check(entries):
 for i in range(1,4):
  part=f'ppt/slides/slide{i}.xml'; slide=ET.fromstring(entries[part]); assert (slide.get('show')=='0') == (i==2)
  layout=ET.fromstring(entries[owner(entries,part,'slideLayout')]); assert layout.find('p:cSld',ns).get('name')==('Target' if i<3 else 'Source'); master=owner(entries,owner(entries,part,'slideLayout'),'slideMaster'); assert master in entries; assert owner(entries,part,'notesSlide') in entries
  shapes=[x for x in slide.findall('.//p:sp',ns) if x.find('p:nvSpPr/p:cNvPr',ns) is not None and x.find('p:nvSpPr/p:cNvPr',ns).get('name')==f'body-{i-1}']; assert len(shapes)==1
  shape=shapes[0]; body=shape.find('p:txBody',ns); run=body.find('a:p/a:r',ns); prop=run.find('a:rPr',ns); text=run.find('a:t',ns).text; assert text==f'Native reset page {i-1}'
  assert prop.get('sz')=='2400' and prop.get('b')==('1' if i==1 else '0'); assert body.find('a:bodyPr',ns).get('wrap')==('square' if i==1 else 'none'); assert body.find('a:p/a:pPr',ns).get('algn')==('l' if i==1 else 'r')
  assert prop.find('a:latin',ns).get('typeface')==('+mn-lt' if i==1 else 'Arial'); paint=prop.find('a:solidFill',ns); node=paint.find('a:schemeClr' if i==1 else 'a:srgbClr',ns); assert node is not None and node.get('val')==('accent1' if i==1 else 'CC2200')
  assert shape.find('p:nvSpPr/p:nvPr/p:ph',ns) is not None
files=['input.pptx','round1.pptx','round2.pptx','round3.pptx']; archives={}
for name in files:
 with zipfile.ZipFile(root/name) as z: archives[name]={n:z.read(n) for n in z.namelist()}
 check(archives[name])
pdf=fitz.open(root/'round3.pdf'); assert len(pdf)==3
spans=[]
for i,page in enumerate(pdf):
 found=[s for b in page.get_text('dict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans'] if s['text']==f'Native reset page {i}']; assert len(found)==1
 s=found[0]; assert s['font']==('CourierNewPS-BoldMT' if i==0 else 'ArialMT') and s['color']==(0x0044cc if i==0 else 0xcc2200) and 24<s['size']<26; spans.append({k:s[k] for k in ('text','font','size','color','bbox')})
# Actual altered ZIP files: each corruption independently crosses a checked invariant.
negative=root/'negatives';negative.mkdir(exist_ok=True); base=archives['round3.pptx']; part='ppt/slides/slide1.xml'
def mutated(edit):
 copy=dict(base); slide=ET.fromstring(copy[part]); shape=next(x for x in slide.findall('.//p:sp',ns) if x.find('p:nvSpPr/p:cNvPr',ns) is not None and x.find('p:nvSpPr/p:cNvPr',ns).get('name')=='body-0');edit(slide,shape);copy[part]=ET.tostring(slide,encoding='utf-8',xml_declaration=True);return copy
edits={'bold':lambda s,x:x.find('.//a:rPr',ns).set('b','0'),'font':lambda s,x:x.find('.//a:rPr/a:latin',ns).set('typeface','Arial'),'color':lambda s,x:x.find('.//a:rPr/a:solidFill/a:schemeClr',ns).set('val','accent2'),'wrap':lambda s,x:x.find('.//a:bodyPr',ns).set('wrap','none'),'align':lambda s,x:x.find('.//a:pPr',ns).set('algn','r'),'text':lambda s,x:setattr(x.find('.//a:t',ns),'text','changed'),'hidden':lambda s,x:s.set('show','0')}
neg={}
for name,edit in edits.items():
 path=negative/(name+'.pptx'); data=mutated(edit)
 with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
  for k,v in data.items():z.writestr(k,v)
 with zipfile.ZipFile(path) as z: reread={n:z.read(n) for n in z.namelist()}
 try: check(reread)
 except (AssertionError,AttributeError,KeyError,ValueError): neg[name]={'sha256':digest(path),'rejected':True}
 else: raise AssertionError('Accepted corruption: '+name)
probe=fitz.open(root/'inheritance-probe.pdf'); assert len(probe)==3
probe_spans=[s for b in probe[0].get_text('dict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans'] if s['text']=='Native reset page 0']; assert len(probe_spans)==1 and probe_spans[0]['font']=='ArialMT' and probe_spans[0]['color']==0
with zipfile.ZipFile(root/'inheritance-probe.pptx') as z:
 probe_slide=ET.fromstring(z.read('ppt/slides/slide1.xml')); probe_shape=next(x for x in probe_slide.findall('.//p:sp',ns) if x.find('p:nvSpPr/p:cNvPr',ns) is not None and x.find('p:nvSpPr/p:cNvPr',ns).get('name')=='body-0'); probe_body=probe_shape.find('p:txBody',ns); probe_run=probe_body.find('a:p/a:r/a:rPr',ns)
 assert 'wrap' not in probe_body.find('a:bodyPr',ns).attrib and 'algn' not in probe_body.find('a:p/a:pPr',ns).attrib and 'b' not in probe_run.attrib and all(probe_run.find('a:'+tag,ns) is None for tag in ('solidFill','latin','ea','cs'))
list_probe=fitz.open(root/'inheritance-list-probe.pdf'); edit_probe=fitz.open(root/'inheritance-list-edit-probe.pdf'); assert len(list_probe)==len(edit_probe)==3
for document,expected_font,expected_color in [(list_probe,'CourierNewPS-BoldMT',0x0044cc),(edit_probe,'Georgia',0x008800)]:
 for i,page in enumerate(document):
  found=[s for b in page.get_text('dict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans'] if s['text']==f'Native reset page {i}']; assert len(found)==1
  assert found[0]['font']==(expected_font if i==0 else 'ArialMT') and found[0]['color']==(expected_color if i==0 else 0xcc2200)
with zipfile.ZipFile(root/'inheritance-list-probe.pptx') as z1, zipfile.ZipFile(root/'inheritance-list-edit-probe.pptx') as z2:
 assert z1.read('ppt/slides/slide1.xml')==z2.read('ppt/slides/slide1.xml')
 for archive,expected_font,expected_bold in [(z1,'+mn-lt','1'),(z2,'Georgia','0')]:
  layout=ET.fromstring(archive.read('ppt/slideLayouts/slideLayout3.xml')); shape=next(x for x in layout.findall('.//p:sp',ns) if x.find('p:nvSpPr/p:cNvPr',ns) is not None and x.find('p:nvSpPr/p:cNvPr',ns).get('name')=='copy'); prop=shape.find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',ns); assert prop.get('b')==expected_bold and prop.find('a:latin',ns).get('typeface')==expected_font
dependencies=[Path(__file__),Path('scripts/prepare-style-reset-native.mjs'),Path('tests/fixtures/layout-rebind.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'producerSourceSha256':{str(p):digest(p) for p in dependencies},'scope':'five selected properties in text control; three saved/reopened PPTX files; final PDF only','nativeApp':'Microsoft PowerPoint 16.104 (25121423)','directory':str(root),'files':{n:digest(root/n) for n in files+['round3.pdf']},'spans':spans,'negativeZip':neg,'omissionProbe':{'pptxSha256':digest(root/'inheritance-probe.pptx'),'pdfSha256':digest(root/'inheritance-probe.pdf'),'observedFont':probe_spans[0]['font'],'observedColor':probe_spans[0]['color'],'result':'Omitting five local properties alone does not reproduce the target blue Courier appearance.'},'layoutListProbe':{'files':{n:digest(root/n) for n in ['inheritance-list-probe.pptx','inheritance-list-probe.pdf','inheritance-list-edit-probe.pptx','inheritance-list-edit-probe.pdf']},'result':'Adding level-one layout defaults restored blue Courier bold; editing only layout defaults produced green Georgia on page 0 while control pages stayed red Arial. Geometry and native UI editing remain unverified.'},'threeRoundVisualEquivalence':False,'otherSelectorsNativeAccepted':False,'futureNativeLayoutInheritanceVerified':False,'observation':'Fresh export writes direct run, paragraph and frame values; live native inheritance after editing the layout is not established.'}
Path('docs/testing/style-reset-native-evidence.json').write_text(json.dumps(report,indent=2)+'\n'); print(json.dumps(report,indent=2))
