"""Check PowerPoint's controlled empty doughnut geometry-attribute sample."""
import hashlib,json,math,plistlib,re,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2]; base=root/'.tmp/native-doughnut-empty-geometry/powerpoint-aPY1ED'; control=root/'.tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pptx'
evidence=Path(next((x for x in sys.argv[1:] if x!='--record'),root/'docs/testing/chart-doughnut-empty-geometry-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest(); manifest=json.loads((base/'manifest.json').read_text());assert sha(control)==manifest['sourceSha256'] and sha(base/'empty.pptx')==manifest['sha256']
changes=[['firstSliceAng'],['holeSize'],['firstSliceAng','holeSize']]; c='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
with zipfile.ZipFile(control) as a,zipfile.ZipFile(base/'empty.pptx') as b:
 assert a.testzip() is None and b.testzip() is None and set(a.namelist())==set(b.namelist())
 for name in a.namelist():
  raw=a.read(name)
  if name in [f'ppt/charts/chart{i}.xml' for i in range(1,4)]:
   i=int(re.search(r'chart(\d+)\.xml',name).group(1))-1
   for tag in changes[i]:
    pattern=rb'<c:'+tag.encode()+rb' val="[0-9]+"/>'; assert len(re.findall(pattern,raw))==1;raw=re.sub(pattern,b'<c:'+tag.encode()+b'/>',raw,count=1)
  assert raw==b.read(name),name
with zipfile.ZipFile(base/'empty-round-1.pptx') as z:
 assert z.testzip() is None
 saved=[]
 for i,expected in enumerate([(0,50),(90,0),(0,0)],1):
  chart=ET.fromstring(z.read(f'ppt/charts/chart{i}.xml')).find('.//'+c+'doughnutChart'); pair=tuple(int(chart.find(c+tag).get('val')) for tag in ('firstSliceAng','holeSize')); assert pair==expected;saved.append({'firstSliceAngle':pair[0],'holeSize':pair[1]})
with fitz.open(base/'empty.pdf') as doc,fitz.open(root/'.tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pdf') as ref:
 assert len(doc)==3
 def colors(page):return {tuple(round(v,3) for v in d['fill']):d for d in page.get_drawings() if d.get('fill') and max(d['fill'])>.5 and min(d['fill'])<.1}
 first=colors(doc[0]); prior=colors(ref[0]); assert first.keys()==prior.keys()
 for color in first: assert all(abs(a-b)<.001 for a,b in zip(first[color]['rect'],prior[color]['rect']))
 red=(1.,0.,0.);blue=(0.,0.,1.);hits=[]
 for i in (1,2):
  ds=colors(doc[i]);box=ds[blue]['rect'] if i==1 else ds[red]['rect'];cx,cy=(box.x0+box.x1)/2,(box.y0+box.y1)/2
  points=[p for item in ds[red]['items'] if item[0]=='l' for p in item[1:]];distance=min(math.hypot(p.x-cx,p.y-cy) for p in points);assert distance<.1;hits.append({'slide':i+1,'center':[cx,cy],'distance':distance})
 captures=[]
 for phase in ('first','reopen'):
  for i,name in enumerate(('rings','rotate','blank'),1):
   ax=base/f'{phase}-{i}.ax.txt';state=ax.read_text();assert f'chart-{name}' in state and f'第 {i} 张幻灯片' in state
   captures.append({'phase':phase,'slide':i,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{phase}-{i}.png')})
result={'scope':'Three Transitional source samples replace existing first-slice angle and/or hole-size values with empty elements. PowerPoint 16.104 opened them; first-print vector geometry matches the explicit zero-angle control on slide 1, while the red paths on slides 2 and 3 reach the center. One Save As wrote explicit zero for missing attribute values; all three slides were captured after closing and reopening the saved file, but the saved file was not reprinted; screenshot and AX hashes do not independently establish provenance. Other producers, versions, Strict files and native data editing are not established.','applicationVersion':plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],'generatorSha256':sha(root/'scripts/prepare-native-doughnut-empty-geometry.mjs'),'verifierSha256':sha(Path(__file__)),'controlSha256':sha(control),'sourceSha256':sha(base/'empty.pptx'),'savedSha256':sha(base/'empty-round-1.pptx'),'pdfSha256':sha(base/'empty.pdf'),'savedGeometry':saved,'centerHits':hits,'captures':captures}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'slides':3,'saveAs':1,'reopenRounds':1}))
