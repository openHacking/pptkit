"""Verify the retained single-attribute PowerPoint repair-warning control."""
import hashlib,json,plistlib,sys,zipfile
from pathlib import Path
root=Path(__file__).resolve().parents[2]; base=root/'.tmp/native-doughnut-percent/powerpoint-JYogRH'; original=root/'.tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pptx'
evidence=Path(next((x for x in sys.argv[1:] if x!='--record'),root/'docs/testing/chart-doughnut-percent-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
manifest=json.loads((base/'manifest.json').read_text()); assert sha(original)==manifest['sourceSha256'] and sha(base/'percent.pptx')==manifest['sha256']
with zipfile.ZipFile(original) as a, zipfile.ZipFile(base/'percent.pptx') as b:
 assert a.testzip() is None and b.testzip() is None and set(a.namelist())==set(b.namelist())
 for name in a.namelist():
  before,after=a.read(name),b.read(name)
  if name=='ppt/charts/chart1.xml':
   old=b'<c:holeSize val="50"/>'; new=b'<c:holeSize val="050%"/>'; assert before.count(old)==1; before=before.replace(old,new)
  assert before==after,name
ax=base/'repair-warning.ax.txt';state=ax.read_text();assert 'percent.pptx' in state and '修复' in state and '取消' in state
result={'scope':'One controlled generated PowerPoint file differs from a previously opened control only in a doughnut hole size lexical value, 50 versus 050%. PowerPoint 16.104 displayed a repair warning; repair was cancelled. Capture hashes do not independently establish provenance. Other lexical forms, OOXML dialects, versions and post-repair behavior were not tested.','applicationVersion':plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],'generatorSha256':sha(root/'scripts/prepare-native-doughnut-percent.mjs'),'verifierSha256':sha(Path(__file__)),'sourceSha256':sha(original),'candidateSha256':sha(base/'percent.pptx'),'axSha256':sha(ax),'screenshotSha256':sha(base/'repair-warning.png')}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else: assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'changedAttributes':1,'powerPointRepairWarning':True}))
