"""Verify PPTKit-authored scatter caches/workbooks and first-open native geometry."""
import hashlib,io,json,math,plistlib,posixpath,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2];base=root/'.tmp/native-scatter-authored/powerpoint-PmQkqF'
evidence=Path(next((a for a in sys.argv[1:] if a!='--record'),root/'docs/testing/chart-scatter-authored-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest();manifest=json.loads((base/'manifest.json').read_text());source=base/'scatter-authored.pptx';layout=json.loads((base/'layout.json').read_text());assert sha(source)==manifest['sha256']==layout['sourceSha256']
c='{http://schemas.openxmlformats.org/drawingml/2006/chart}';s='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';r='{http://schemas.openxmlformats.org/package/2006/relationships}'
expected=[([-8,None,2,8],[4,3,-4,6],'A','B'),([7,-1,3,-7],[-6,4,2,-2],'C','D')];modes=['gap','gap','zero','span'];packages=[]
for phase,file in [('source',source),('saved',base/'scatter-authored-round-1.pptx')]:
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None;charts=[]
  for i,mode in enumerate(modes,1):
   xml=ET.fromstring(z.read(f'ppt/charts/chart{i}.xml'));chart=xml.find('.//'+c+'scatterChart');assert chart is not None and xml.find('.//'+c+'dispBlanksAs').get('val')==mode
   axes=xml.findall('.//'+c+'valAx');assert len(axes)==2 and not xml.findall('.//'+c+'catAx')
   for axis in axes:
    scaling=axis.find(c+'scaling');assert float(scaling.find(c+'min').get('val'))==-10 and float(scaling.find(c+'max').get('val'))==10 and float(axis.find(c+'majorUnit').get('val'))==5
   series=chart.findall(c+'ser');assert len(series)==2
   rel=ET.fromstring(z.read(f'ppt/charts/_rels/chart{i}.xml.rels'));targets=[x.get('Target') for x in rel.findall(r+'Relationship') if x.get('Type','').endswith('/package')];assert len(targets)==1;bookpart=posixpath.normpath(posixpath.join('ppt/charts',targets[0]))
   with zipfile.ZipFile(io.BytesIO(z.read(bookpart))) as book:
    assert book.testzip() is None;sheet=ET.fromstring(book.read('xl/worksheets/sheet1.xml'));cells={cell.get('r'):cell for cell in sheet.iter(s+'c')}
    for node,(xs,ys,xc,yc) in zip(series,expected):
     for tag,col,values in [('xVal',xc,xs),('yVal',yc,ys)]:
      ref=node.find(c+tag+'/'+c+'numRef');assert ref.find(c+'f').text==f'Sheet1!${col}$2:${col}$5';cache=ref.find(c+'numCache');assert cache.find(c+'ptCount').get('val')=='4'
      assert [(int(p.get('idx')),float(p.find(c+'v').text)) for p in cache.findall(c+'pt')]==[(j,v) for j,v in enumerate(values) if v is not None]
      for row,value in enumerate(values,2):
       v=cells[f'{col}{row}'].find(s+'v');assert (v is None if value is None else v is not None and float(v.text)==value)
   charts.append({'mode':mode,'workbookPart':bookpart})
  packages.append({'phase':phase,'sha256':sha(file),'charts':charts})
comparisons=[]
with fitz.open(base/'scatter-authored.pdf') as pdf:
 assert len(pdf)==4
 for slide,(page,model) in enumerate(zip(pdf,layout['slides']),1):
  frame=next(d['rect'] for d in page.get_drawings() if d.get('fill')==(1.,1.,1.) and d['rect'].width>700);sx,sy=frame.width/960,frame.height/540;entries=[]
  for j,color in enumerate([(1.,0.,0.),(0.,0.,1.)]):
   markers=[d for d in page.get_drawings() if d.get('fill')==color];points=model['scatter'][j]['points'];assert len(markers)==len(points)
   def convert(x,y):return ((x-frame.x0)/sx,(y-frame.y0)/sy)
   def error(native,point):return math.hypot(native[0]-model['box']['x']-point['x'],native[1]-model['box']['y']-point['y'])
   markerErrors=[error(convert((d['rect'].x0+d['rect'].x1)/2,(d['rect'].y0+d['rect'].y1)/2),point) for d,point in zip(markers,points)];assert max(markerErrors)<3.2
   lines=[item for d in page.get_drawings() if d.get('color')==color and d.get('fill') is None for item in d['items'] if item[0]=='l']
   segments=[] if slide==1 else [(a,b) for segment in model['scatter'][j]['segments'] for a,b in zip(segment,segment[1:])];assert len(lines)==len(segments)
   lineErrors=[error(convert(point.x,point.y),target) for line,(a,b) in zip(lines,segments) for point,target in [(line[1],a),(line[2],b)]];assert max(lineErrors,default=0)<3.2
   entries.append({'series':j,'markers':len(points),'segments':len(segments),'maxMarkerDifference':max(markerErrors),'maxLineDifference':max(lineErrors,default=0)})
  comparisons.append({'slide':slide,'series':entries})
captures=[]
for phase in ('first','reopen'):
 for i,name in enumerate(['markers-gap','line-gap','line-zero','line-span'],1):
  ax=base/f'{phase}-{i}.ax.txt';state=ax.read_text();assert f'chart-{name}' in state and f'第 {i} 张幻灯片' in state
  captures.append({'phase':phase,'slide':i,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{phase}-{i}.png')})
result={'scope':'Four PPTKit-authored fixed-axis scatter charts with missing X and unsorted series were opened in PowerPoint 16.104, printed, saved and reopened once. Source and saved caches, formulas, workbook cells and axes were checked. First-print marker centers and line endpoints were compared with the authored Layout snapshot. The samples show an approximately 2.5–2.8 slide-point horizontal difference, so this is approximate geometry, not pixel parity. The snapshot is reconstructed by the inspector, not imported from the source; captures do not independently prove provenance. Scatter content import, automatic axes, interactive editing, other producers and reopened PDF remain unverified.','applicationVersion':plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],'generatorSha256':sha(root/'scripts/prepare-native-scatter-authored.mjs'),'inspectorSha256':sha(root/'scripts/inspect-native-scatter-authored-layout.mjs'),'verifierSha256':sha(Path(__file__)),'layoutSha256':sha(base/'layout.json'),'pdfSha256':sha(base/'scatter-authored.pdf'),'packages':packages,'comparisons':comparisons,'captures':captures}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'charts':4,'reopenRounds':1,'maxMarkerDifference':max(x['maxMarkerDifference'] for c in comparisons for x in c['series']),'maxLineDifference':max(x['maxLineDifference'] for c in comparisons for x in c['series'])}))
