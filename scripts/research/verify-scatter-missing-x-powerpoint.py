"""Check independent scatter data, saved package, native markers and paths."""
import hashlib, io, json, math, plistlib, posixpath, sys, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import fitz

root=Path(__file__).resolve().parents[2]
base=root/'.tmp/native-scatter-missing-x/powerpoint-tp3m_1ti'
evidence=Path(next((a for a in sys.argv[1:] if a!='--record'),root/'docs/testing/chart-scatter-missing-x-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
manifest=json.loads((base/'manifest.json').read_text());source=base/'scatter-missing-x.pptx';assert sha(source)==manifest['sha256']
c='{http://schemas.openxmlformats.org/drawingml/2006/chart}';a='{http://schemas.openxmlformats.org/drawingml/2006/main}';s='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';r='{http://schemas.openxmlformats.org/package/2006/relationships}'
modes=['gap','gap','zero','span']; names=manifest['groups'];groups=[('RED',[(-8,4),(None,3),(2,-4),(8,6)]),('BLUE',[(7,-6),(-1,4),(3,2),(-7,-2)])]
packages=[]
for phase,file in [('source',source),('saved',base/'scatter-missing-x-round-1.pptx')]:
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None;charts=[]
  for i,mode in enumerate(modes,1):
   xml=ET.fromstring(z.read(f'ppt/charts/chart{i}.xml'));chart=xml.find('.//'+c+'scatterChart');assert chart is not None
   assert xml.find('.//'+c+'dispBlanksAs').get('val')==mode
   axes=xml.findall('.//'+c+'valAx');assert len(axes)==2
   for axis in axes:
    sc=axis.find(c+'scaling');assert float(sc.find(c+'min').get('val'))==-10 and float(sc.find(c+'max').get('val'))==10;assert float(axis.find(c+'majorUnit').get('val'))==5
   assert chart.find(c+'scatterStyle').get('val')=='lineMarker'
   ser=chart.findall(c+'ser');assert len(ser)==2
   for node,(_,pairs) in zip(ser,groups):
    line=node.find(c+'spPr/'+a+'ln');assert line is not None and line.find(a+('noFill' if i==1 else 'solidFill')) is not None
    for tag,values in [('xVal',[x for x,y in pairs]),('yVal',[y for x,y in pairs])]:
     cache=node.find(c+tag+'/'+c+'numRef/'+c+'numCache');assert cache.find(c+'ptCount').get('val')=='4'
     assert [(int(p.get('idx')),float(p.find(c+'v').text)) for p in cache.findall(c+'pt')]==[(j,v) for j,v in enumerate(values) if v is not None]
   rels=ET.fromstring(z.read(f'ppt/charts/_rels/chart{i}.xml.rels'));targets=[x.get('Target') for x in rels.findall(r+'Relationship') if x.get('Type','').endswith('/package')];assert len(targets)==1
   bookpart=posixpath.normpath(posixpath.join('ppt/charts',targets[0]))
   with zipfile.ZipFile(io.BytesIO(z.read(bookpart))) as book:
    assert book.testzip() is None;sheet=ET.fromstring(book.read('xl/worksheets/sheet1.xml'));cells={x.get('r'):x for x in sheet.iter(s+'c')}
    for first,(_,pairs) in zip([2,8],groups):
     for j,(x,y) in enumerate(pairs,first):
      vx=cells[f'A{j}'].find(s+'v');assert (vx is None if x is None else vx is not None and float(vx.text)==x);v=cells[f'B{j}'].find(s+'v');assert (v is None if y is None else v is not None and float(v.text)==y)
   charts.append({'mode':mode,'workbookPart':bookpart})
  packages.append({'phase':phase,'sha256':sha(file),'charts':charts})

def fit(values,observed):
 mean_x=sum(values)/len(values);mean_y=sum(observed)/len(observed)
 slope=sum((x-mean_x)*(y-mean_y) for x,y in zip(values,observed))/sum((x-mean_x)**2 for x in values)
 intercept=mean_y-slope*mean_x
 return intercept,slope

samples=[]
with fitz.open(base/'scatter-missing-x.pdf') as pdf:
 assert len(pdf)==4
 for page_index,page in enumerate(pdf):
  mode=modes[page_index];observed=[]
  for color in [(1.,0.,0.),(0.,0.,1.)]:
   markers=[d for d in page.get_drawings() if d.get('fill')==color]
   observed.append([[(d['rect'].x0+d['rect'].x1)/2,(d['rect'].y0+d['rect'].y1)/2] for d in markers])
  expected=[[(0 if x is None else x,0 if y is None else y) for x,y in pairs if (x is not None and y is not None) or mode=='zero'] for _,pairs in groups]
  assert [len(x) for x in observed]==[len(x) for x in expected]
  pairs=[(point,actual) for group,actuals in zip(expected,observed) for point,actual in zip(group,actuals)]
  ix,sx=fit([p[0][0] for p in pairs],[p[1][0] for p in pairs]);iy,sy=fit([p[0][1] for p in pairs],[p[1][1] for p in pairs]);assert sx>0 and sy<0
  errors=[math.hypot(ix+sx*p[0]-q[0],iy+sy*p[1]-q[1]) for p,q in pairs];assert max(errors)<.2
  line_counts=[];max_endpoint_difference=0
  for j,color in enumerate([(1.,0.,0.),(0.,0.,1.)]):
   paths=[d for d in page.get_drawings() if d.get('color')==color and d.get('fill') is None]
   segments=[item for path in paths for item in path['items']];assert all(t[0]=='l' for t in segments)
   red={0:[],1:[(1,2)],2:[(0,1),(1,2),(2,3)],3:[(0,1),(1,2)]}[page_index]
   blue=[] if page_index==0 else [(0,1),(1,2),(2,3)]
   chosen=red if j==0 else blue;assert len(segments)==len(chosen)
   for item,(a,b) in zip(segments,chosen):
    for endpoint,index in [(item[1],a),(item[2],b)]:
     center=observed[j][index];error=math.hypot(endpoint.x-center[0],endpoint.y-center[1]);assert error<.3;max_endpoint_difference=max(max_endpoint_difference,error)
   line_counts.append(len(segments))
  samples.append({'slide':page_index+1,'mode':mode,'markerCounts':[len(x) for x in observed],'markers':observed,'lineCounts':line_counts,'fit':{'xIntercept':ix,'xPerUnit':sx,'yIntercept':iy,'yPerUnit':sy},'maxMarkerResidual':max(errors),'maxEndpointDifference':max_endpoint_difference})
captures=[]
for phase in ('first','reopen'):
 for i,name in enumerate(names,1):
  ax=base/f'{phase}-{i}.ax.txt';state=ax.read_text();assert f'chart-{name}' in state and f'第 {i} 张幻灯片' in state
  captures.append({'phase':phase,'slide':i,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{phase}-{i}.png')})
result={'scope':'Independent python-pptx XY chart/workbook source with explicit gap, zero or span scalar. Four fixed-axis samples opened in PowerPoint 16.104, printed once, then saved and reopened once. Marker centers were fitted to the input X/Y values, and native line endpoints and missing-X behavior and unsorted-series order were checked. Source and saved caches and workbooks were checked. This is not PPTKit scatter output; captures do not independently prove provenance. Other producers, other missing-X positions and both-coordinates-blank, interactive edits, automatic axes and reopened PDF remain unverified.','applicationVersion':plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],'generatorSha256':sha(root/'scripts/prepare-native-scatter-missing-x.py'),'verifierSha256':sha(Path(__file__)),'pdfSha256':sha(base/'scatter-missing-x.pdf'),'packages':packages,'samples':samples,'captures':captures}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'charts':4,'reopenRounds':1,'maxMarkerResidual':max(x['maxMarkerResidual'] for x in samples),'maxEndpointDifference':max(x['maxEndpointDifference'] for x in samples)}))
