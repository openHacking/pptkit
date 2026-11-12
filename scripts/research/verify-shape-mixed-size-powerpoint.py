"""Measure first-open PowerPoint mixed-size shape tabs against guides and SVG's baseline estimate."""
import hashlib
import json
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
import fitz
root = Path(__file__).resolve().parents[2]
base = root/'.tmp/native-tab-stops/powerpoint-shape-mixed-size-774YI9'
evidence = Path(next((arg for arg in sys.argv[1:] if arg != '--record'), root/'docs/testing/tab-shape-mixed-size-powerpoint-evidence.json'))
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
source, pdf = base/'shapes.pptx', base/'shapes.pdf'
svg_source = root/'packages/svg-renderer/src/rendering/text.ts'
assert 'const tabY = mixedSizes || eastAsian ? y + line.fontSize * (1 - TEXT_BASELINE_RATIO) : y;' in svg_source.read_text()
assert sha(source) == json.loads((base/'manifest.json').read_text())['sha256']
a='{http://schemas.openxmlformats.org/drawingml/2006/main}'; p='{http://schemas.openxmlformats.org/presentationml/2006/main}'
cases=[('left','l',['I','WIDE','MIXED']),('center','ctr',['I','WIDE','MIXED']),('right','r',['I','WIDE','MIXED']),('decimal','dec',['1.2','123.45','0.678'])]
with zipfile.ZipFile(source) as package:
 assert package.testzip() is None
 for page,(kind,align,words) in enumerate(cases,1):
  slide=ET.fromstring(package.read(f'ppt/slides/slide{page}.xml'))
  shapes={s.find(p+'nvSpPr/'+p+'cNvPr').get('name'):s for s in slide.iter(p+'sp')}
  for i,word in enumerate(words):
   para=shapes[f'{kind}-{i}'].find(p+'txBody/'+a+'p'); runs=para.findall(a+'r')
   assert [r.find(a+'t').text for r in runs] == ['P\t',word]
   assert [r.find(a+'rPr').get('sz') for r in runs] == ['2000',str([2800,3600,1600][i])]
   assert para.find(a+'pPr/'+a+'tabLst/'+a+'tab').attrib == {'pos':'3556000','algn':align}
records=[]
with fitz.open(pdf) as doc:
 assert len(doc)==4
 for page,(kind,_,words) in zip(doc,cases):
  drawings=page.get_drawings(); fills=[d for d in drawings if d['fill'] and all(abs(d['fill'][i]-v/255)<.01 for i,v in enumerate((242,245,250)))]; assert len(fills)==3
  guides=[d for d in drawings if d['fill'] and d['fill'][0]>.7 and d['fill'][1]<.3]; assert len(guides)==1
  guide=(guides[0]['rect'].x0+guides[0]['rect'].x1)/2
  spans=[s for b in page.get_text('dict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans']]
  rawspans=[s for b in page.get_text('rawdict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans']]
  dots=[c['origin'][0] for s in rawspans for c in s['chars'] if c['c']=='.']
  rows=[]
  for i,word in enumerate(words):
   target=[s for s in spans if s['text']==word]; assert len(target)==1
   s=target[0]; box=s['bbox']; anchor=box[0] if kind=='left' else (box[0]+box[2])/2 if kind=='center' else box[2] if kind=='right' else dots[i]
   assert abs(anchor-guide)<.5
   scale=fills[i]['rect'].height/65; baseline=(s['origin'][1]-fills[i]['rect'].y0)/scale
   estimate=max(20,[28,36,16][i]); delta=baseline-estimate
   assert abs(delta) < 1
   xscale=fills[i]['rect'].width/700
   characters = next(s['chars'] for s in rawspans if ''.join(c['c'] for c in s['chars']) == word)
   relative_chars = [{'char':c['c'], 'left':(c['bbox'][0]-fills[i]['rect'].x0)/xscale, 'right':(c['bbox'][2]-fills[i]['rect'].x0)/xscale, 'baseline':(c['origin'][1]-fills[i]['rect'].y0)/scale} for c in characters]
   rows.append({'characters':relative_chars, 'word':word,'suffixSize':[28,36,16][i], 'anchor':round(anchor,4), 'powerpointBaselineFromTopPt':round(baseline,4), 'svgEstimatedBaselineFromTopPt':estimate, 'baselineDifferencePt':round(delta,4)})
  records.append({'kind':kind,'guide':round(guide,4),'rows':rows})
state=(base/'first.ax.txt').read_text(); assert 'shapes.pptx' in state and 'Help: left-1, Value: P\tWIDE' in state
version=plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString']
result={'scope':'First-open PowerPoint A4 print of 12 mixed-size shape text samples. Measures guide alignment and compares PDF baseline against the SVG renderer formula, not a full browser or visual-equivalence test. Only the first page window was captured; editing, reopening, other fonts and transforms remain unverified. Capture hashes do not establish provenance independently.', 'applicationVersion':version, 'sourceSha256':sha(source), 'pdfSha256':sha(pdf), 'axSha256':sha(base/'first.ax.txt'), 'screenshotSha256':sha(base/'first.png'), 'generatorSha256':sha(root/'scripts/prepare-native-shape-tab-mixed-size.mjs'), 'svgSourceSha256':sha(svg_source), 'verifierSha256':sha(Path(__file__)), 'pages':records}
if '--record' in sys.argv: evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else: assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'samples':12,'baselineDifferenceRangePt':[min(r['baselineDifferencePt'] for p in records for r in p['rows']),max(r['baselineDifferencePt'] for p in records for r in p['rows'])]}))
