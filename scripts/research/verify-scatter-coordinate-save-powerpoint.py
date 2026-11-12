"""Verify one native scatter-coordinate preservation sample and retained captures."""
import hashlib
import io
import json
import plistlib
import posixpath
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
BASE = ROOT / '.tmp/native-scatter-coordinate-save/powerpoint-bI3pt5'
RECORD = ROOT / 'docs/testing/chart-scatter-coordinate-save-powerpoint-evidence.json'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/package/2006/relationships}'

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def inspect(filename, expected):
    with ZipFile(BASE / filename) as package:
        chart = ET.fromstring(package.read('ppt/charts/chart1.xml'))
        series = chart.findall('.//' + C + 'scatterChart/' + C + 'ser')
        assert len(series) == 2
        rels = ET.fromstring(package.read('ppt/charts/_rels/chart1.xml.rels'))
        rel = [r for r in rels.findall(R + 'Relationship') if r.get('Type', '').endswith('/package')]
        assert len(rel) == 1 and rel[0].get('TargetMode', 'Internal') == 'Internal'
        workbook = posixpath.normpath(posixpath.join('ppt/charts', rel[0].get('Target')))
        with ZipFile(io.BytesIO(package.read(workbook))) as book:
            sheet = ET.fromstring(book.read('xl/worksheets/sheet1.xml'))
            cells = {cell.get('r'): cell.findtext(S + 'v') for cell in sheet.findall('.//' + S + 'c')}
        for index, (node, columns) in enumerate(zip(series, [('A', 'B'), ('C', 'D')])):
            for dimension, column in zip(('xVal', 'yVal'), columns):
                ref = node.find(C + dimension + '/' + C + 'numRef')
                assert ref is not None and ref.findtext(C + 'f') == f'Sheet1!${column}$2:${column}$4'
                cache = ref.find(C + 'numCache')
                assert cache is not None and cache.find(C + 'ptCount').get('val') == '3'
                points = cache.findall(C + 'pt')
                assert [p.get('idx') for p in points] == ['0', '1', '2']
                values = [float(p.findtext(C + 'v')) for p in points]
                wanted = [point[0 if dimension == 'xVal' else 1] for point in expected[index]]
                assert values == wanted
                assert [float(cells[f'{column}{row}']) for row in (2, 3, 4)] == wanted
    return {'file': str((BASE / filename).relative_to(ROOT)), 'sha256': sha(BASE / filename)}

before = [[[-8, 4], [2, -4], [8, 6]], [[7, -6], [-1, 4], [3, 2]]]
after = [[[-3, 8], [2, -4], [8, 6]], [[7, -6], [-1, -7], [3, 2]]]
packages = [inspect('source.pptx', before), inspect('edited.pptx', after), inspect('native-round-1.pptx', after)]
manifest = json.loads((BASE / 'manifest.json').read_text())
assert manifest['sourceSha256'] == packages[0]['sha256'] and manifest['editedSha256'] == packages[1]['sha256']
assert manifest['expected'] == {'RED': after[0], 'BLUE': after[1]}
assert sorted((v['coordinate'], v['before'], v['after']) for v in manifest['changes']) == [('x', -8, -3), ('y', 4, -7), ('y', 4, 8)]
with ZipFile(BASE / 'source.pptx') as source, ZipFile(BASE / 'edited.pptx') as edited:
    assert set(source.namelist()) == set(edited.namelist())
    changed = {name for name in source.namelist() if source.read(name) != edited.read(name)}
    assert changed == {'ppt/charts/chart1.xml', 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'}
captures = []
for phase, title in [('first', 'edited'), ('reopen', 'native-round-1')]:
    ax = (BASE / f'{phase}.ax.txt').read_text()
    assert f'Window: "{title}"' in ax and 'Scatter coordinate preservation' in ax
    for pair in ('(-3, 8)', '(2, -4)', '(8, 6)', '(7, -6)', '(-1, -7)', '(3, 2)'):
        assert pair in ax
    captures.append({'phase': phase, 'axSha256': sha(BASE / f'{phase}.ax.txt'), 'screenshotSha256': sha(BASE / f'{phase}.png')})
result = {'scope': 'One dense two-series fixed-axis PPTKit-authored scatter chart had three source-bound X/Y coordinate edits. PowerPoint opened the edited file, saved a copy, closed and reopened it once. Cache values, workbook cells, formulas, the source-to-edited changed parts and the two retained accessibility captures were checked. The captures are retained observations, not independently authenticated provenance. Native interactive workbook editing, sparse coordinates, other producers, geometry measurements and further reopen rounds are not established.',
    'applicationVersion': plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],
    'generatorSha256': sha(ROOT / 'scripts/prepare-native-scatter-coordinate-save.mjs'), 'verifierSha256': sha(Path(__file__)),
    'manifestSha256': sha(BASE / 'manifest.json'), 'packages': packages, 'captures': captures}
if sys.argv[1:] == ['--record']:
    RECORD.write_text(json.dumps(result, indent=2) + '\n')
else:
    record = Path(sys.argv[1]) if len(sys.argv) == 2 else RECORD
    assert json.loads(record.read_text()) == result
print(json.dumps({'verified': True, 'charts': 1, 'coordinateEdits': 3, 'reopenRounds': 1}))
