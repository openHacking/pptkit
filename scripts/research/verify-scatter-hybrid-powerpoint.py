"""Verify PPTKit-generated XY series/workbooks inside an independent chart shell."""
import hashlib,io,json,math,plistlib,posixpath,re,subprocess,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2];base=root/'.tmp/native-scatter-hybrid/powerpoint-V2fSwr'
evidence=Path(next((a for a in sys.argv[1:] if a!='--record'),root/'docs/testing/chart-scatter-hybrid-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest();manifest=json.loads((base/'manifest.json').read_text());source=base/'scatter-hybrid.pptx';template=root/manifest['source']
assert sha(source)==manifest['sha256'] and sha(template)==manifest['sourceSha256']
subprocess.run(['python3',str(root/'scripts/research/verify-scatter-missing-x-powerpoint.py')],check=True,capture_output=True)
c='{http://schemas.openxmlformats.org/drawingml/2006/chart}';s='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';r='{http://schemas.openxmlformats.org/package/2006/relationships}'
expected=[([-8,None,2,8],[4,3,-4,6],'A','B'),([7,-1,3,-7],[-6,4,2,-2],'C','D')];modes=['gap','gap','zero','span'];packages=[]
with zipfile.ZipFile(template) as original,zipfile.ZipFile(source) as hybrid:
 assert set(original.namelist())==set(hybrid.namelist())
 for part in original.namelist():
  if part.startswith('ppt/embeddings/'):continue
  if re.fullmatch(r'ppt/charts/chart[1-4]\.xml',part):
   strip=lambda b:re.sub(rb'<c:ser>.*?</c:ser>',b'',b,flags=re.S)
   assert strip(original.read(part))==strip(hybrid.read(part))
  else:assert original.read(part)==hybrid.read(part),part
for phase,file in [('source',source),('saved',base/'scatter-hybrid-round-1.pptx')]:
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None;charts=[]
  for i,mode in enumerate(modes,1):
   xml=ET.fromstring(z.read(f'ppt/charts/chart{i}.xml'));chart=xml.find('.//'+c+'scatterChart');assert chart is not None and xml.find('.//'+c+'dispBlanksAs').get('val')==mode
   series=chart.findall(c+'ser');assert len(series)==2
   rel=ET.fromstring(z.read(f'ppt/charts/_rels/chart{i}.xml.rels'));targets=[x.get('Target') for x in rel.findall(r+'Relationship') if x.get('Type','').endswith('/package')];assert len(targets)==1
   bookpart=posixpath.normpath(posixpath.join('ppt/charts',targets[0]))
   with zipfile.ZipFile(io.BytesIO(z.read(bookpart))) as book:
    assert book.testzip() is None;sheet=ET.fromstring(book.read('xl/worksheets/sheet1.xml'));cells={cell.get('r'):cell for cell in sheet.iter(s+'c')}
    for node,(xs,ys,xc,yc) in zip(series,expected):
     for tag,col,values in [('xVal',xc,xs),('yVal',yc,ys)]:
      ref=node.find(c+tag+'/'+c+'numRef');assert ref.find(c+'f').text==f'Sheet1!${col}$2:${col}$5'
      cache=ref.find(c+'numCache');assert cache.find(c+'ptCount').get('val')=='4'
      assert [(int(p.get('idx')),float(p.find(c+'v').text)) for p in cache.findall(c+'pt')]==[(j,v) for j,v in enumerate(values) if v is not None]
      for row,value in enumerate(values,2):
       v=cells[f'{col}{row}'].find(s+'v');assert (v is None if value is None else v is not None and float(v.text)==value)
   charts.append({'mode':mode,'workbookPart':bookpart})
  packages.append({'phase':phase,'sha256':sha(file),'charts':charts})
comparisons=[]
with fitz.open(root/'.tmp/native-scatter-missing-x/powerpoint-tp3m_1ti/scatter-missing-x.pdf') as control,fitz.open(base/'scatter-hybrid.pdf') as actual:
 assert len(control)==len(actual)==4
 for slide,(a,b) in enumerate(zip(control,actual),1):
  results=[]
  for color in [(1.,0.,0.),(0.,0.,1.)]:
   def shapes(page):return [d for d in page.get_drawings() if d.get('fill')==color]
   def lines(page):return [item for d in page.get_drawings() if d.get('color')==color and d.get('fill') is None for item in d['items'] if item[0]=='l']
   am,bm=shapes(a),shapes(b);al,bl=lines(a),lines(b);assert len(am)==len(bm) and len(al)==len(bl)
   marker=max((max(abs(x-y) for x,y in zip(xd['rect'],yd['rect'])) for xd,yd in zip(am,bm)),default=0)
   line=max((math.hypot(x[j].x-y[j].x,x[j].y-y[j].y) for x,y in zip(al,bl) for j in (1,2)),default=0)
   assert marker<.01 and line<.01
   results.append({'color':list(color),'markers':len(am),'segments':len(al),'maxMarkerBoxDifference':marker,'maxLineEndpointDifference':line})
  comparisons.append({'slide':slide,'series':results})
captures=[]
for phase in ('first','reopen'):
 for i,name in enumerate(['markers-gap','line-gap','line-zero','line-span'],1):
  ax=base/f'{phase}-{i}.ax.txt';state=ax.read_text();assert f'chart-{name}' in state and f'第 {i} 张幻灯片' in state
  captures.append({'phase':phase,'slide':i,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{phase}-{i}.png')})
result={'scope':'An independently generated fixed-axis XY chart shell was retained byte-for-byte except for series XML and embedded workbooks replaced by PPTKit internal builders. Four samples with missing X and unsorted data opened in PowerPoint 16.104, printed, then saved and reopened once. Source and saved caches, formulas and workbook cells were checked, and first-open marker boxes and line endpoints were compared with the independent native control. Captures do not independently prove provenance. This is a hybrid research artifact, not public PPTKit scatter export; native interactive editing, other producers, automatic axes and reopened PDF remain unverified.','applicationVersion':plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],'generatorSha256':sha(root/'scripts/prepare-native-scatter-hybrid.mjs'),'verifierSha256':sha(Path(__file__)),'seriesBuilderSha256':sha(root/'packages/pptx-exporter/src/ooxml/scatter-series.ts'),'workbookBuilderSha256':sha(root/'packages/pptx-exporter/src/ooxml/scatter-workbook.ts'),'referenceBuilderSha256':sha(root/'packages/pptx-exporter/src/ooxml/chart-reference.ts'),'worksheetPackagingSha256':sha(root/'packages/pptx-exporter/src/ooxml/chart-workbook.ts'),'chartStringSha256':sha(root/'packages/pptx-exporter/src/ooxml/chart-string.ts'),'xmlWriterSha256':sha(root/'packages/pptx-exporter/src/ooxml/xml.ts'),'templateSha256':sha(template),'pdfSha256':sha(base/'scatter-hybrid.pdf'),'controlPdfSha256':sha(root/'.tmp/native-scatter-missing-x/powerpoint-tp3m_1ti/scatter-missing-x.pdf'),'packages':packages,'comparisons':comparisons,'captures':captures}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'charts':4,'reopenRounds':1,'maxMarkerBoxDifference':max(v['maxMarkerBoxDifference'] for c in comparisons for v in c['series']),'maxLineEndpointDifference':max(v['maxLineEndpointDifference'] for c in comparisons for v in c['series'])}))
