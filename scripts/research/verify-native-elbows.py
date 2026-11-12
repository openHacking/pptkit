"""Verify retained PowerPoint first-open and reopened PDF outlines for free elbows."""
import hashlib
import json
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
import fitz

root = Path(__file__).resolve().parents[2]
base = root / '.tmp/native-connectors/powerpoint-elbow-6OzKx0'
evidence = Path(next((arg for arg in sys.argv[1:] if arg != '--record'), root / 'docs/testing/native-elbow-powerpoint-evidence.json'))
sha = lambda file: hashlib.sha256(file.read_bytes()).hexdigest()
source, pdf = base/'elbows.pptx', base/'elbows.pdf'
manifest = json.loads((base/'manifest.json').read_text()); assert sha(source) == manifest['sourceSha256']
p = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
directions = [(1,1),(-1,1),(1,-1),(-1,-1)]
def verify_package(file):
 with zipfile.ZipFile(file) as package:
     assert package.testzip() is None
     for index, kind in enumerate(('single','double'), 1):
         slide = ET.fromstring(package.read(f'ppt/slides/slide{index}.xml'))
         connectors = {item.find(p+'nvCxnSpPr/'+p+'cNvPr').get('name'): item for item in slide.iter(p+'cxnSp')}
         assert len(connectors) == 4
         for i, (sx, sy) in enumerate(directions):
             node = connectors[f'{kind}-{i}']; geometry = node.find(p+'spPr/'+a+'prstGeom')
             assert geometry.get('prst') == ('bentConnector2' if kind == 'single' else 'bentConnector3')
             if kind == 'double': assert geometry.find(a+'avLst/'+a+'gd').attrib == {'name':'adj1','fmla':'val 25000'}
             xfrm = node.find(p+'spPr/'+a+'xfrm')
             assert (xfrm.get('flipH') == '1') == (sx < 0) and (xfrm.get('flipV') == '1') == (sy < 0)
             assert xfrm.find(a+'ext').attrib == {'cx':'2540000','cy':'1270000'}
             cx, cy = (680 if i % 2 else 220), (160 if i < 2 else 360)
             assert xfrm.find(a+'off').attrib == {'x':str((cx-100)*12700),'y':str((cy-50)*12700)}
             line = node.find(p+'spPr/'+a+'ln'); assert line.get('w') == '38100'
             assert line.find(a+'headEnd').get('type') == 'oval' and line.find(a+'tailEnd').get('type') == 'triangle'
verify_package(source)
saved = base/'elbows-round-1.pptx'
verify_package(saved)
observed = []
for phase, printed in (('first', pdf), ('reopen', base/'elbows-round-1.pdf')):
 with fitz.open(printed) as document:
     assert len(document) == 2
     for page_index, (page, kind) in enumerate(zip(document, ('single','double')), 1):
         drawings = page.get_drawings(); white = [d for d in drawings if d['fill'] == (1.0, 1.0, 1.0) and d['rect'].width > 700]
         assert len(white) >= 1
         rect = max(white, key=lambda d: d['rect'].width)['rect']; xscale, yscale = rect.width/960, rect.height/540
         red = [d for d in drawings if d['fill'] and d['fill'][0] > .7 and d['fill'][1] < .3]; assert len(red) == 4
         records = []
         for i, ((sx, sy), item) in enumerate(zip(directions, red)):
             cx, cy = (680 if i % 2 else 220), (160 if i < 2 else 360)
             startx, starty, endx, endy = cx-sx*100, cy-sy*50, cx+sx*100, cy+sy*50
             bendx = endx if kind == 'single' else startx+(endx-startx)*.25
             X = lambda value: rect.x0 + value*xscale
             Y = lambda value: rect.y0 + value*yscale
             first, vertical = item['items'][0], item['items'][2]
             assert first[0] == 'l' and vertical[0] == 'l'
             expected = [X(startx), Y(starty)-sy*1.5*yscale, X(bendx), X(bendx)+sx*1.5*xscale, Y(starty)]
             actual = [first[1].x, first[1].y, first[2].x, vertical[1].x, vertical[1].y]
             assert all(abs(a-b) < .05 for a,b in zip(actual, expected)), (kind, i, actual, expected)
             if kind == 'double':
                 line = item['items'][4]; assert line[0] == 'l'
                 assert abs(line[1].x-X(bendx)) < .05 and abs(line[1].y-(Y(endy)-sy*1.5*yscale)) < .05
             records.append({'name': f'{kind}-{i}', 'outline': [round(v,4) for v in actual], 'expected': [round(v,4) for v in expected]})
         ax = base/f'{phase}-{page_index}.ax.txt'; state = ax.read_text()
         assert f'{kind}-title' in state and all(f'形状 (disabled) {kind}-{i}\n' in state for i in range(4))
         observed.append({'phase':phase,'kind':kind,'pageRect':[round(v,4) for v in rect], 'samples':records, 'axSha256':sha(ax), 'screenshotSha256':sha(base/f'{phase}-{page_index}.png')})
version = plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString']
result = {'scope':'PowerPoint first-open and one save/close/reopen with A4 scaled prints: four free single elbows and four free double elbows with 25% adjustment. Checks selected PDF stroke-outline coordinates and source XML; no attached targets, interactive editing, additional reopen rounds or general preset claim. Capture hashes do not independently prove provenance.', 'applicationVersion':version, 'directory':str(base.relative_to(root)), 'sourceSha256':sha(source), 'pdfSha256':sha(pdf), 'savedSha256':sha(saved), 'reopenedPdfSha256':sha(base/'elbows-round-1.pdf'), 'generatorSha256':sha(root/'scripts/prepare-native-elbow-controls.mjs'), 'verifierSha256':sha(Path(__file__)), 'pages':observed}
if '--record' in sys.argv: evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else: assert json.loads(evidence.read_text()) == result
print(json.dumps({'verified':True,'pages':len(observed),'samples':sum(len(page['samples']) for page in observed)}))
