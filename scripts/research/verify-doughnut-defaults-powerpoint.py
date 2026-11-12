"""Verify omitted doughnut geometry against PowerPoint's first print and save."""
import hashlib,json,math,plistlib,re,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2]
base=root/'.tmp/native-doughnut-defaults/powerpoint-FE4zPv'
original=root/'.tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pptx'
evidence=Path(next((x for x in sys.argv[1:] if x!='--record'),root/'docs/testing/chart-doughnut-defaults-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
manifest=json.loads((base/'manifest.json').read_text()); assert sha(original)==manifest['sourceSha256'] and sha(base/'defaults.pptx')==manifest['sha256']
c='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
removed=[['firstSliceAng'],['holeSize'],['firstSliceAng','holeSize']]
with zipfile.ZipFile(original) as src, zipfile.ZipFile(base/'defaults.pptx') as case:
 assert src.testzip() is None and case.testzip() is None and set(src.namelist())==set(case.namelist())
 for name in src.namelist():
  a,b=src.read(name),case.read(name)
  if name.startswith('ppt/charts/chart') and name.endswith('.xml') and name in [f'ppt/charts/chart{i}.xml' for i in range(1,4)]:
   i=int(re.search(r'chart(\d+)\.xml',name).group(1))-1
   for item in removed[i]: a=re.sub(rb'<c:'+item.encode()+rb' val="[0-9]+"/>',b'',a,count=1)
  assert a==b, name
packages=[]
for phase,file,expected in [('source',base/'defaults.pptx',[(None,50),(90,None),(None,None)]),('saved',base/'defaults-round-1.pptx',[(0,50),(90,0),(0,0)])]:
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None; rows=[]
  for i,(angle,hole) in enumerate(expected,1):
   xml=ET.fromstring(z.read(f'ppt/charts/chart{i}.xml')); chart=xml.find('.//'+c+'doughnutChart'); assert chart is not None
   def val(tag):
    item=chart.find(c+tag); return None if item is None else int(item.get('val'))
   assert (val('firstSliceAng'),val('holeSize'))==(angle,hole)
   rows.append({'firstSliceAngle':angle,'holeSize':hole})
  packages.append({'phase':phase,'sha256':sha(file),'geometry':rows})
with fitz.open(base/'defaults.pdf') as doc, fitz.open(root/'.tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pdf') as control:
 assert len(doc)==3 and len(control)==3
 def colors(page): return {tuple(round(v,3) for v in d['fill']):d for d in page.get_drawings() if d.get('fill') and max(d['fill'])>.5 and min(d['fill'])<.1}
 first=colors(doc[0]); prior=colors(control[0]); assert first.keys()==prior.keys()
 for color in first:
  assert all(abs(a-b)<.001 for a,b in zip(first[color]['rect'],prior[color]['rect']))
 red=(1.,0.,0.); blue=(0.,0.,1.); center_hits=[]
 for i in (1,2):
  drawings=colors(doc[i]); box=drawings[blue]['rect'] if i==1 else drawings[red]['rect']; cx,cy=(box.x0+box.x1)/2,(box.y0+box.y1)/2
  red_path=drawings[red]['items']; ends=[p for item in red_path if item[0]=='l' for p in item[1:]]
  distance=min(math.hypot(p.x-cx,p.y-cy) for p in ends); assert distance<.1
  center_hits.append({'slide':i+1,'center':[cx,cy],'nearestRedPathPointDistance':distance})
 captures=[]
 for phase in ('first','reopen'):
  for i,name in enumerate(('rings','rotate','blank'),1):
   ax=base/f'{phase}-{i}.ax.txt'; state=ax.read_text(); assert f'chart-{name}' in state and f'第 {i} 张幻灯片' in state
   captures.append({'phase':phase,'slide':i,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{phase}-{i}.png')})
result={'scope':'Three controlled variants remove only doughnut geometry elements. After first-open captures and one save/close/reopen of a copy, the original was reopened and reprinted; that PDF matches the explicit zero-angle control for slide 1; slide 2 and 3 red paths reach the center when holeSize is absent. One saved/reopened copy writes zero for the missing hole size and angle. These observations do not establish other producers, later Office versions, empty val attributes or native data edits; captures do not independently prove provenance.','applicationVersion':plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],'generatorSha256':sha(root/'scripts/prepare-native-doughnut-defaults.mjs'),'verifierSha256':sha(Path(__file__)),'controlPdfSha256':sha(root/'.tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pdf'),'pdfSha256':sha(base/'defaults.pdf'),'packages':packages,'centerHits':center_hits,'captures':captures}
if '--record' in sys.argv: evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else: assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'slides':3,'reopenRounds':1,'missingHoleResavedAs':0}))
