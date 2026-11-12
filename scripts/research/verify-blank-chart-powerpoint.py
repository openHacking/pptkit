"""Verify retained first-open and three-reopen PowerPoint blank-chart observations."""
import hashlib,io,json,plistlib,posixpath,re,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
root=Path(__file__).resolve().parents[2];base=root/'.tmp/native-chart-blanks/powerpoint-AP8EmX'
evidence=Path(next((x for x in sys.argv[1:] if x!='--record'),root/'docs/testing/chart-blank-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
source=base/'blank-charts.pptx';saved=[base/f'blank-charts-round-{i}.pptx' for i in range(1,4)];assert sha(source)==json.loads((base/'manifest.json').read_text())['sha256']
c='{http://schemas.openxmlformats.org/drawingml/2006/chart}';s='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';rel='{http://schemas.openxmlformats.org/package/2006/relationships}'
packages=[]
for file in (source,*saved):
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None; parts=[]
  for i,kind in enumerate(('bar','line','pie'),1):
   name=f'ppt/charts/chart{i}.xml';chart=ET.fromstring(z.read(name));assert chart.find('.//'+c+'dispBlanksAs').get('val')=='gap'
   cache=chart.find('.//'+c+'ser/'+c+'val/'+c+'numRef/'+c+'numCache');assert cache.find(c+'ptCount').get('val')=='4'
   assert [(v.get('idx'),v.find(c+'v').text) for v in cache.findall(c+'pt')]==[('0','4'),('2','0'),('3','3')]
   relationships=ET.fromstring(z.read(f'ppt/charts/_rels/chart{i}.xml.rels'));targets=[r.get('Target') for r in relationships.findall(rel+'Relationship') if r.get('Type','').endswith('/package')];assert len(targets)==1
   bookpart=posixpath.normpath(posixpath.join('ppt/charts',targets[0]));
   with zipfile.ZipFile(io.BytesIO(z.read(bookpart))) as book:
    assert book.testzip() is None; sheet=ET.fromstring(book.read('xl/worksheets/sheet1.xml'))
    cells={cell.get('r'):cell for cell in sheet.iter(s+'c')};assert cells['B3'].find(s+'v') is None
    assert [cells[f'B{row}'].find(s+'v').text for row in (2,4,5)]==['4','0','3']
   parts.append({'kind':kind,'chartPart':name,'workbookPart':bookpart})
  packages.append({'sha256':sha(file),'charts':parts})
captures=[]
for phase in ('first','reopen','round-2','round-3'):
 for i,kind in enumerate(('bar','line','pie'),1):
  label=f'{phase}-page-{i}' if phase.startswith('round-') else f'{phase}-{i}'
  ax=base/f'{label}.ax.txt';state=ax.read_text();assert f'chart-{kind}' in state and f'第 {i} 张幻灯片' in state
  points=dict(re.findall(r'点 "([^"]+)"\s*\n值: ?([^\n]*)',state))
  assert points.get('A','').startswith('4') and points.get('ZERO','').startswith('0') and points.get('D','').startswith('3')
  assert points.get('BLANK','')==''
  captures.append({'phase':phase,'kind':kind,'points':points,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{label}.png')})
version=plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString']
result={'scope':'PowerPoint first open and three sequential save/close/reopen rounds for one bar, one line and one pie chart with [4, blank, 0, 3]. Checks cache and embedded workbook through relationships plus AX point values on all pages; line screenshot visually shows a gap. No in-app data edit, other distributions or general Office parity claim. Capture hashes do not independently prove provenance.', 'applicationVersion':version,'generatorSha256':sha(root/'scripts/prepare-native-blank-charts.mjs'),'verifierSha256':sha(Path(__file__)),'packages':packages,'captures':captures}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'charts':3,'captureStates':len(captures),'reopenRounds':3}))
