"""Independently check PowerPoint small doughnut holes and saved data."""
import hashlib,io,json,math,plistlib,posixpath,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2];base=root/'.tmp/native-doughnut-small-hole/powerpoint-Kx43uN'
evidence=Path(next((x for x in sys.argv[1:] if x!='--record'),root/'docs/testing/chart-doughnut-small-hole-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest(); source=base/'small-holes.pptx';saved=base/'small-holes-round-1.pptx';manifest=json.loads((base/'manifest.json').read_text());assert sha(source)==manifest['sha256']
warnings=manifest['warnings'];assert [w['elementId'] for w in warnings if w['code']=='doughnut-small-hole-office-extension']==['chart-1','chart-5','chart-9']
c='{http://schemas.openxmlformats.org/drawingml/2006/chart}';s='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';r='{http://schemas.openxmlformats.org/package/2006/relationships}'
packages=[]
for phase,file in [('source',source),('saved',saved)]:
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None; charts=[]
  for i,size in enumerate((1,5,9,10),1):
   part=f'ppt/charts/chart{i}.xml';xml=ET.fromstring(z.read(part));assert int(xml.find('.//'+c+'holeSize').get('val'))==size
   cache=xml.find('.//'+c+'ser/'+c+'val/'+c+'numRef/'+c+'numCache');assert cache.find(c+'ptCount').get('val')=='2' and [(p.get('idx'),p.find(c+'v').text) for p in cache.findall(c+'pt')]==[('0','1'),('1','3')]
   rels=ET.fromstring(z.read(f'ppt/charts/_rels/chart{i}.xml.rels')); targets=[x.get('Target') for x in rels.findall(r+'Relationship') if x.get('Type','').endswith('/package')];assert len(targets)==1
   bookpart=posixpath.normpath(posixpath.join('ppt/charts',targets[0]))
   with zipfile.ZipFile(io.BytesIO(z.read(bookpart))) as book:
    assert book.testzip() is None; sheet=ET.fromstring(book.read('xl/worksheets/sheet1.xml')); cells={x.get('r'):x for x in sheet.iter(s+'c')};assert [cells[k].find(s+'v').text for k in ('B2','B3')]==['1','3']
   charts.append({'holeSize':size,'chartPart':part,'workbookPart':bookpart})
  packages.append({'phase':phase,'sha256':sha(file),'charts':charts})
with fitz.open(base/'small-holes.pdf') as doc:
 assert len(doc)==4;measurements=[]
 for page,size in zip(doc,(1,5,9,10)):
  red=[d for d in page.get_drawings() if d.get('fill')==(1.,0.,0.)];assert len(red)==1; d=red[0];rect=d['rect']; cx,cy=rect.x0,rect.y1;outer=cy-rect.y0;line=d['items'][-1];assert line[0]=='l';inner=math.hypot(line[1].x-cx,line[1].y-cy);ratio=inner/outer;assert abs(ratio-size/100)<.0001
  measurements.append({'holeSize':size,'nativeInnerOuterRadiusRatio':ratio,'outerRect':list(rect)})
captures=[]
for phase in ('first','reopen'):
 for i,size in enumerate((1,5,9,10),1):
  ax=base/f'{phase}-{i}.ax.txt';state=ax.read_text();assert f'chart-{size}' in state and f'第 {i} 张幻灯片' in state
  captures.append({'phase':phase,'slide':i,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{phase}-{i}.png')})
result={'scope':'Four generated 1%, 5%, 9% and 10% charts observed on first open and one save/close/reopen in PowerPoint 16.104. The first-open PDF vector paths establish the four hole/outer radius ratios. Source and saved packages retain cache points and workbook values. Screenshots and AX hashes do not independently establish provenance; reopened PDF, native data editing, other applications and versions are not established.','applicationVersion':plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],'generatorSha256':sha(root/'scripts/prepare-native-doughnut-small-hole.mjs'),'verifierSha256':sha(Path(__file__)),'packages':packages,'firstPdfSha256':sha(base/'small-holes.pdf'),'measurements':measurements,'captures':captures}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'charts':4,'reopenRounds':1}))
