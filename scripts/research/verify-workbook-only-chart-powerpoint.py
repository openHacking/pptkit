"""Bind first-open PowerPoint observations for three charts without numeric caches."""
import hashlib,json,plistlib,re,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
root=Path(__file__).resolve().parents[2];base=root/'.tmp/native-chart-blanks/powerpoint-AP8EmX'
evidence=Path(next((x for x in sys.argv[1:] if x!='--record'),root/'docs/testing/chart-workbook-only-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
source,fixture=base/'blank-charts.pptx',base/'blank-workbook-only.pptx'
assert sha(source)==json.loads((base/'manifest.json').read_text())['sha256']
c='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
with zipfile.ZipFile(source) as before,zipfile.ZipFile(fixture) as after:
 assert before.testzip() is None and after.testzip() is None and set(before.namelist())==set(after.namelist())
 chartnames={f'ppt/charts/chart{i}.xml' for i in range(1,4)}
 for name in before.namelist():
  original,modified=before.read(name),after.read(name)
  if name in chartnames:
   expected,n=re.subn(rb'<c:numCache>[\s\S]*?</c:numCache>',b'',original);assert n==1 and modified==expected
   tree=ET.fromstring(modified);assert tree.find('.//'+c+'ser/'+c+'val/'+c+'numRef') is not None and tree.find('.//'+c+'ser/'+c+'val/'+c+'numRef/'+c+'numCache') is None
  else:assert original==modified
captures=[]
for i,kind in enumerate(('bar','line','pie'),1):
 ax=base/f'workbook-only-{i}.ax.txt';state=ax.read_text();assert 'blank-workbook-only.pptx' in state and f'chart-{kind}' in state and f'第 {i} 张幻灯片' in state
 assert '点 "' not in state and '绘图区' not in state
 captures.append({'kind':kind,'axSha256':sha(ax),'screenshotSha256':sha(base/f'workbook-only-{i}.png')})
version=plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString']
result={'scope':'First-open PowerPoint observation of three charts with numeric caches removed and otherwise unchanged embedded workbooks. All three AX trees expose chart objects but no plot area or data points; the first-page screenshot was visually inspected and shows no chart. Excel was not installed. No editing, refresh, print, save/reopen or other-environment claim; capture hashes do not independently prove provenance.', 'applicationVersion':version,'originalSha256':sha(source),'fixtureSha256':sha(fixture),'generatorSha256':sha(root/'scripts/prepare-native-workbook-only-chart.mjs'),'verifierSha256':sha(Path(__file__)),'captures':captures}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'charts':3,'numericCachesRemoved':3}))
