"""Verify native blank display strategies against saved XML, workbook and PDF paths."""
import hashlib,io,json,math,plistlib,posixpath,re,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2];base=root/'.tmp/native-chart-blanks/powerpoint-modes-vkUEcD'
evidence=Path(next((x for x in sys.argv[1:] if x!='--record'),root/'docs/testing/chart-blank-modes-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
source,saved=base/'blank-modes.pptx',base/'blank-modes-round-1.pptx';assert sha(source)==json.loads((base/'manifest.json').read_text())['sha256']
c='{http://schemas.openxmlformats.org/drawingml/2006/chart}';s='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';rel='{http://schemas.openxmlformats.org/package/2006/relationships}'
cases=[('line','gap'),('line','zero'),('line','span'),('bar','zero'),('pie','zero')]
packages=[]
for file in (source,saved):
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None; charts=[]
  for i,(family,mode) in enumerate(cases,1):
   part=f'ppt/charts/chart{i}.xml';xml=ET.fromstring(z.read(part));assert xml.find('.//'+c+'dispBlanksAs').get('val')==mode
   cache=xml.find('.//'+c+'ser/'+c+'val/'+c+'numRef/'+c+'numCache');assert cache.find(c+'ptCount').get('val')=='5'
   assert [(v.get('idx'),v.find(c+'v').text) for v in cache.findall(c+'pt')]==[('0','4'),('2','1'),('4','3')]
   relationships=ET.fromstring(z.read(f'ppt/charts/_rels/chart{i}.xml.rels'));targets=[r.get('Target') for r in relationships.findall(rel+'Relationship') if r.get('Type','').endswith('/package')];assert len(targets)==1
   bookpart=posixpath.normpath(posixpath.join('ppt/charts',targets[0]))
   with zipfile.ZipFile(io.BytesIO(z.read(bookpart))) as book:
    assert book.testzip() is None;sheet=ET.fromstring(book.read('xl/worksheets/sheet1.xml'));cells={v.get('r'):v for v in sheet.iter(s+'c')}
    assert all(cells[f'B{row}'].find(s+'v') is None for row in (3,5))
    assert [cells[f'B{row}'].find(s+'v').text for row in (2,4,6)]==['4','1','3']
   charts.append({'family':family,'mode':mode,'chartPart':part,'workbookPart':bookpart})
  packages.append({'sha256':sha(file),'charts':charts})
observed=[]
for phase,pdf in (('first',base/'blank-modes.pdf'),('reopen',base/'blank-modes-round-1.pdf')):
 with fitz.open(pdf) as doc:
  assert len(doc)==5;pages=[]
  for page,mode in zip(doc,('gap','zero','span')):
   blue=[d for d in page.get_drawings() if d['color'] and .2<d['color'][0]<.4 and .35<d['color'][1]<.55 and .65<d['color'][2]<.85]
   markers=[(d['rect'].x0+d['rect'].width/2,d['rect'].y0+d['rect'].height/2) for d in blue if d['width']<.7 and d['rect'].x0<650 and 4<d['rect'].width<6]
   paths=[d for d in blue if d['width']>1 and any(v[0]=='l' and abs(v[2].x-v[1].x)>50 for v in d['items'])]
   assert len(markers)==(5 if mode=='zero' else 3) and len(paths)==(0 if mode=='gap' else 1)
   segments=[] if not paths else [(tuple(v[1]),tuple(v[2])) for v in paths[0]['items'] if v[0]=='l'];assert len(segments)=={'gap':0,'zero':4,'span':2}[mode]
   for (start,end),left,right in zip(segments,markers,markers[1:]):
    assert math.dist(start,left)<.5 and math.dist(end,right)<.5
   pages.append({'mode':mode,'markers':[[round(x,4),round(y,4)] for x,y in markers],'segments':[[[round(v,4) for v in p] for p in edge] for edge in segments]})
  z=pages[1]['markers'];assert all(abs((z[i][0]-z[0][0])-(z[4][0]-z[0][0])*i/4)<.5 for i in range(5))
  unit=z[1][1]-z[2][1];assert unit>0 and abs(z[0][1]-(z[1][1]-4*unit))<1 and abs(z[4][1]-(z[1][1]-3*unit))<1 and abs(z[3][1]-z[1][1])<.5
  for page in (pages[0],pages[2]):
   for actual,expected in zip(page['markers'],[z[0],z[2],z[4]]):assert math.dist(actual,expected)<.5
  observed.append({'phase':phase,'pdfSha256':sha(pdf),'linePages':pages})
captures=[]
for phase in ('first','reopen'):
 for i,(family,mode) in enumerate(cases,1):
  ax=base/f'{phase}-{i}.ax.txt';state=ax.read_text();assert f'chart-{family}-{mode}' in state and f'第 {i} 张幻灯片' in state
  points=dict(re.findall(r'点 "([^"]+)"\s*\n值: ?([^\n]*)',state));assert points.get('A','').startswith('4') and points.get('C','').startswith('1') and points.get('E','').startswith('3')
  assert all(points.get(name,'').strip() in ('','()') for name in ('BLANK1','BLANK2'))
  captures.append({'phase':phase,'family':family,'mode':mode,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{phase}-{i}.png')})
version=plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString']
result={'scope':'PowerPoint first-open and one save/close/reopen for line gap/zero/span and bar/pie zero. Verifies all source/saved caches, workbooks and AX states, and independently measures the three line policies in both A4 PDFs. No interactive data edit, other datasets or general application parity claim. Capture hashes do not independently prove provenance.', 'applicationVersion':version,'generatorSha256':sha(root/'scripts/prepare-native-chart-blank-modes.mjs'),'verifierSha256':sha(Path(__file__)),'packages':packages,'pdfs':observed,'captures':captures}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'charts':5,'linePolicies':3,'pdfs':2,'reopenRounds':1}))
