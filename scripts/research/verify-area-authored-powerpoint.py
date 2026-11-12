"""Check a controlled area-chart research package and native PDF polygons."""
import hashlib,io,json,math,plistlib,posixpath,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2];base=root/'.tmp/native-area-authored/powerpoint-ohuZBk'
evidence=Path(next((a for a in sys.argv[1:] if a!='--record'),root/'docs/testing/chart-area-authored-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest(); source=base/'area-authored.pptx';manifest=json.loads((base/'manifest.json').read_text());assert sha(source)==manifest['sha256'];layout=json.loads((base/'layout.json').read_text());assert layout['sourceSha256']==sha(source)
c='{http://schemas.openxmlformats.org/drawingml/2006/chart}';s='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';r='{http://schemas.openxmlformats.org/package/2006/relationships}'
expected=[('standard','gap'),('standard','zero'),('stacked','zero'),('percentStacked','zero')];packages=[]
for phase,file in [('source',source)]+[(f'saved-{n}',base/f'area-authored-round-{n}.pptx') for n in (1,2,3)]:
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None;charts=[]
  for i,(group,mode) in enumerate(expected,1):
   xml=ET.fromstring(z.read(f'ppt/charts/chart{i}.xml'));area=xml.find('.//'+c+'areaChart');assert area.find(c+'grouping').get('val')==group and xml.find('.//'+c+'dispBlanksAs').get('val')==mode
   series=area.findall(c+'ser');assert len(series)==2
   for ser,values in zip(series,[[4,2,None,-2,3],[2,-3,3,4,1]]):
    cache=ser.find('.//'+c+'val/'+c+'numRef/'+c+'numCache');assert cache.find(c+'ptCount').get('val')=='5';assert [(int(x.get('idx')),float(x.find(c+'v').text)) for x in cache.findall(c+'pt')]==[(j,v) for j,v in enumerate(values) if v is not None]
   rels=ET.fromstring(z.read(f'ppt/charts/_rels/chart{i}.xml.rels'));targets=[x.get('Target') for x in rels.findall(r+'Relationship') if x.get('Type','').endswith('/package')];assert len(targets)==1;bookpart=posixpath.normpath(posixpath.join('ppt/charts',targets[0]))
   with zipfile.ZipFile(io.BytesIO(z.read(bookpart))) as book:
    assert book.testzip() is None; sheet=ET.fromstring(book.read('xl/worksheets/sheet1.xml'));cells={x.get('r'):x for x in sheet.iter(s+'c')};assert cells['B4'].find(s+'v') is None
    for col,values in [('B',[4,2,None,-2,3]),('C',[2,-3,3,4,1])]:
     for row,value in enumerate(values,2):
      cell=cells[f'{col}{row}'];v=cell.find(s+'v');assert (v is None if value is None else v is not None and float(v.text)==value)
    assert [cells[f'A{k}'].find('.//'+s+'t').text for k in range(2,7)]==list('ABCDE')
   charts.append({'grouping':group,'blankDisplay':mode,'workbookPart':bookpart})
  packages.append({'phase':phase,'sha256':sha(file),'charts':charts})
comparisons=[];all_vertices=[]
for phase,pdf_name in [('first','area-authored.pdf'),('reopen3','area-authored-round-3.pdf')]:
 with fitz.open(base/pdf_name) as pdf:
  assert len(pdf)==4;vertices=[]
  assert all(label in pdf[3].get_text() for label in ['-100%', '-50%', '0%', '50%', '100%'])
  for slide,(page,model,(group,mode)) in enumerate(zip(pdf,layout['slides'],expected),1):
   assert (model['grouping'],model['blankDisplay'])==(group,mode)
   frame=next(d['rect'] for d in page.get_drawings() if d.get('fill')==(1.,1.,1.) and d['rect'].width>700);sx,sy=frame.width/960,frame.height/540; entries=[]
   for series,color in enumerate([(1.,0.,0.),(0.,0.,1.)]):
    paths=sorted([d for d in page.get_drawings() if d.get('fill')==color],key=lambda d:d['rect'].x0);polygons=model['areas'][series]['polygons'];assert len(paths)==len(polygons)
    for path,polygon in zip(paths,polygons):
     items=path['items'];assert all(x[0]=='l' for x in items);points=[items[0][1]]+[x[2] for x in items];points=points[:-1];assert len(points)==len(polygon)
     native=[((a.x-frame.x0)/sx,(a.y-frame.y0)/sy) for a in points];vertices.extend(native)
     errors=[math.hypot(a[0]-(model['box']['x']+b['x']),a[1]-(model['box']['y']+b['y'])) for a,b in zip(native,polygon)];assert max(errors)<.8
     entries.append({'series':series,'vertices':len(points),'maxSlidePointDifference':max(errors)})
   comparisons.append({'phase':phase,'slide':slide,'grouping':group,'blankDisplay':mode,'polygons':entries})
  all_vertices.append(vertices)
assert len(all_vertices[0])==len(all_vertices[1]);reopen_delta=max(math.hypot(a[0]-b[0],a[1]-b[1]) for a,b in zip(*all_vertices));assert reopen_delta<.01
captures=[]
for phase in ('first','reopen','reopen2','reopen3'):
 for i,name in enumerate(['standard-gap','standard-zero','stacked-zero','percentStacked-zero'],1):
  ax=base/f'{phase}-{i}.ax.txt';state=ax.read_text();assert f'chart-{name}' in state and f'第 {i} 张幻灯片' in state
  captures.append({'phase':phase,'slide':i,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{phase}-{i}.png')})
result={'scope':'Four PPTKit-authored area charts with fixed axes and mixed signed values were opened in PowerPoint 16.104 and saved and reopened three times. First-open and third-reopen PDF paths were compared vertex-by-vertex with Layout for standard gap/zero and signed stacked/percent-stacked zero, and with each other; percentage labels were checked. Source and all saved caches and workbooks were checked. Captures do not independently prove provenance; native data editing, other producers and general style parity remain unverified.','applicationVersion':plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],'generatorSha256':sha(root/'scripts/prepare-native-area-authored.mjs'),'layoutInspectorSha256':sha(root/'scripts/inspect-native-area-authored-layout.mjs'),'verifierSha256':sha(Path(__file__)),'layoutSha256':sha(base/'layout.json'),'firstPdfSha256':sha(base/'area-authored.pdf'),'reopen3PdfSha256':sha(base/'area-authored-round-3.pdf'),'maxReopenVertexDifference':reopen_delta,'packages':packages,'comparisons':comparisons,'captures':captures}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'charts':4,'reopenRounds':3,'maxReopenVertexDifference':reopen_delta,'maxSlidePointDifference':max(p['maxSlidePointDifference'] for c in comparisons for p in c['polygons'])}))
