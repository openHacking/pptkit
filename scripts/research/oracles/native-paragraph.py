#!/usr/bin/env python3
"""Independent scoped native checks for seven paragraph default properties."""
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import fitz
root=Path(sys.argv[1]).resolve();P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'p':P,'a':A}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def unpack(p):
 with zipfile.ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
def shape(x,name):
 rows=[s for s in x.findall('.//p:sp',ns) if s.find('p:nvSpPr/p:cNvPr',ns) is not None and s.find('p:nvSpPr/p:cNvPr',ns).get('name')==name];assert len(rows)==1;return rows[0]
def check(d):
 layout=E.fromstring(d['ppt/slideLayouts/slideLayout3.xml']);assert layout.find('p:cSld',ns).get('name')=='Target';lvl=shape(layout,'copy').find('p:txBody/a:lstStyle/a:lvl1pPr',ns);assert lvl.attrib=={'algn':'r','marL':'342900','indent':'-114300'}
 assert lvl.find('a:lnSpc/a:spcPct',ns).get('val')=='150000' and lvl.find('a:spcBef/a:spcPts',ns).get('val')=='300' and lvl.find('a:spcAft/a:spcPts',ns).get('val')=='500' and lvl.find('a:buChar',ns).get('char')=='*' and len(lvl)==4
 for i in range(1,4):
  slide=E.fromstring(d[f'ppt/slides/slide{i}.xml']);assert (slide.get('show')=='0')==(i==2);s=shape(slide,f'body-{i-1}');para=s.find('p:txBody/a:p',ns);assert para.find('a:r/a:t',ns).text==f'Paragraph native page {i-1}';r=para.find('a:r/a:rPr',ns);assert r.find('a:latin',ns).get('typeface')=='Arial' and r.find('a:solidFill/a:srgbClr',ns).get('val')=='CC2200';pr=para.find('a:pPr',ns)
  if i==1:assert pr is None or (not pr.attrib and len(pr)==0)
  else:assert pr is not None and pr.get('algn')=='l' and pr.get('marL')=='0' and int(pr.get('indent'))==0 and pr.find('a:buNone',ns) is not None
files=['inherit.pptx','round1.pptx','round2.pptx','round3.pptx'];data={n:unpack(root/n) for n in files}
for d in data.values():check(d)
pdfs=[fitz.open(root/n) for n in ('materialize.pdf','inherit.pdf','round3.pdf')];assert all(len(d)==3 for d in pdfs)
for i in range(3):
 pages=[d[i] for d in pdfs];assert pages[0].get_pixmap(alpha=False).samples==pages[1].get_pixmap(alpha=False).samples==pages[2].get_pixmap(alpha=False).samples
 for method in ('get_texttrace','get_drawings'):
  v=[getattr(p,method)() for p in pages];assert v[0]==v[1]==v[2]
 v=[p.get_image_info(hashes=True) for p in pages];assert v[0]==v[1]==v[2]
edit=unpack(root/'layout-edit.pptx');assert all(edit[f'ppt/slides/slide{i}.xml']==data['inherit.pptx'][f'ppt/slides/slide{i}.xml'] for i in range(1,4));changed=fitz.open(root/'layout-edit.pdf');assert len(changed)==3
for i,p in enumerate(changed):
 assert (p.get_pixmap(alpha=False).samples==pdfs[0][i].get_pixmap(alpha=False).samples)==(i!=0)
 original=pdfs[1][i].get_text();now=p.get_text();assert f'Paragraph native page {i}' in now
 if i==0:assert '*' in original and '*' not in now
negative=root/'negative';negative.mkdir(exist_ok=True);results={}
for name in ('align','indent','hanging','line','before','after','bullet','layout','hidden'):
 d=dict(data['round3.pptx']);part='ppt/slideLayouts/slideLayout3.xml' if name=='layout' else 'ppt/slides/slide1.xml';x=E.fromstring(d[part])
 if name=='hidden':x.set('show','0')
 elif name=='layout':shape(x,'copy').find('p:txBody/a:lstStyle/a:lvl1pPr',ns).set('algn','l')
 else:
  para=shape(x,'body-0').find('p:txBody/a:p',ns);pr=para.find('a:pPr',ns)
  if pr is None:pr=E.Element('{'+A+'}pPr');para.insert(0,pr)
  if name in ('align','indent','hanging'):pr.set({'align':'algn','indent':'marL','hanging':'indent'}[name],'0')
  elif name=='bullet':E.SubElement(pr,'{'+A+'}buNone')
  else:
   tag={'line':'lnSpc','before':'spcBef','after':'spcAft'}[name];child=E.SubElement(pr,'{'+A+'}'+tag);E.SubElement(child,'{'+A+'}'+('spcPct' if name=='line' else 'spcPts'),{'val':'0'})
 d[part]=E.tostring(x,encoding='utf-8',xml_declaration=True);p=negative/(name+'.pptx')
 with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
  for k,v in d.items():z.writestr(k,v)
 try:check(unpack(p))
 except (AssertionError,AttributeError,KeyError,StopIteration):results[name]={'rejected':True,'sha256':sha(p)}
 else:raise AssertionError('Accepted '+name)
paths=files+['materialize.pptx','layout-edit.pptx','materialize.pdf','inherit.pdf','round3.pdf','layout-edit.pdf'];sources=[Path(__file__),Path('scripts/prepare-native-paragraph.mjs'),Path('tests/fixtures/layout-rebind.mjs'),Path('pnpm-lock.yaml')]+sorted(Path('packages/core/src').rglob('*.ts'))+sorted(Path('packages/layout/src').rglob('*.ts'))+sorted(Path('packages/pptx-exporter/src').rglob('*.ts'))
report={'scope':'named seven-property paragraph control; three native saved/reopened packages, first and final PDF','directory':str(root),'files':{f:sha(root/f) for f in paths},'sourceSha256':{str(p):sha(p) for p in sources},'pixelsTextTraceDrawingImagesEqual':True,'layoutEditChangedSelectedOnly':True,'nativeDropsEmptyParagraphProperties':True,'negativeZip':results,'individualNativePropertyEffectsVerified':False,'nativeInteractiveEditingVerified':False,'sharedFontMetricsResolved':False};Path('docs/testing/native-paragraph-evidence.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
