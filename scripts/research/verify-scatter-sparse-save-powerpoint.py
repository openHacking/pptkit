"""Validate sparse coordinate saves against a retained PowerPoint reopen."""
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
BASE = ROOT / '.tmp/native-scatter-sparse-save/powerpoint-CbeAdy'
RECORD = ROOT / 'docs/testing/chart-scatter-sparse-save-powerpoint-evidence.json'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R = '{http://schemas.openxmlformats.org/package/2006/relationships}'
MODES = ['gap', 'gap', 'zero', 'span']
BEFORE = [[[-8, 4], [-3, None], [2, -4], [8, 6]], [[-7, -2], [-1, 4], [3, 2], [7, -6]]]
AFTER = [[[-6, 4], [-3, None], [2, -3], [8, 6]], [[-7, -2], [-1, 4], [3, 2], [7, -6]]]

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def inspect(filename, groups):
    books = []
    with ZipFile(BASE / filename) as package:
        assert package.testzip() is None
        for index, mode in enumerate(MODES, 1):
            root = ET.fromstring(package.read(f'ppt/charts/chart{index}.xml'))
            assert root.find('.//' + C + 'dispBlanksAs').get('val') == mode
            chart = root.find('.//' + C + 'scatterChart')
            series = chart.findall(C + 'ser')
            assert len(series) == 2
            axes = root.findall('.//' + C + 'valAx')
            assert len(axes) == 2
            ids = [axis.find(C + 'axId').get('val') for axis in axes]
            assert ids == ['-2128940872', '-2129643912']
            rels = ET.fromstring(package.read(f'ppt/charts/_rels/chart{index}.xml.rels'))
            targets = [rel.get('Target') for rel in rels.findall(R + 'Relationship') if rel.get('Type', '').endswith('/package')]
            assert len(targets) == 1
            workbook = posixpath.normpath(posixpath.join('ppt/charts', targets[0]))
            books.append(workbook)
            with ZipFile(io.BytesIO(package.read(workbook))) as book:
                assert book.testzip() is None
                sheet = ET.fromstring(book.read('xl/worksheets/sheet1.xml'))
                cells = {cell.get('r'): cell for cell in sheet.iter(S + 'c')}
            for node, pairs, first in zip(series, groups, [2, 8]):
                for tag, coordinate, col in [('xVal', 0, 'A'), ('yVal', 1, 'B')]:
                    ref = node.find(C + tag + '/' + C + 'numRef')
                    assert ref is not None
                    assert ref.findtext(C + 'f') == f'Sheet1!${col}${first}:${col}${first + 3}'
                    cache = ref.find(C + 'numCache')
                    assert cache is not None and cache.find(C + 'ptCount').get('val') == '4'
                    expected = [pair[coordinate] for pair in pairs]
                    assert [(int(point.get('idx')), float(point.findtext(C + 'v'))) for point in cache.findall(C + 'pt')] == [(i, value) for i, value in enumerate(expected) if value is not None]
                    for offset, value in enumerate(expected):
                        cell = cells.get(f'{col}{first + offset}')
                        actual = cell.find(S + 'v') if cell is not None else None
                        assert actual is None if value is None else actual is not None and float(actual.text) == value
    return {'phase': filename, 'sha256': sha(BASE / filename), 'workbookParts': books}

packages = [inspect('source.pptx', BEFORE), inspect('edited.pptx', AFTER), inspect('native-round-1.pptx', AFTER)]
manifest = json.loads((BASE / 'manifest.json').read_text())
assert manifest['sourceSha256'] == packages[0]['sha256'] and manifest['editedSha256'] == packages[1]['sha256']
assert manifest['expected'] == {'RED': AFTER[0], 'BLUE': AFTER[1]}
assert sorted((change['coordinate'], change['before'], change['after']) for change in manifest['changes']) == [('x', -8, -6)] * 4 + [('y', -4, -3)] * 4
with ZipFile(BASE / 'source.pptx') as source, ZipFile(BASE / 'edited.pptx') as edited:
    assert set(source.namelist()) == set(edited.namelist())
    changed = {name for name in source.namelist() if source.read(name) != edited.read(name)}
    assert changed == {f'ppt/charts/chart{i}.xml' for i in range(1, 5)} | set(packages[0]['workbookParts'])
captures = []
for phase, title in [('first', 'edited'), ('reopen', 'native-round-1')]:
    for index, name in enumerate(['markers-gap', 'line-gap', 'line-zero', 'line-span'], 1):
        ax = BASE / f'{phase}-{index}.ax.txt'
        image = BASE / f'{phase}-{index}.jpg'
        text = ax.read_text()
        assert f'Window: "{title}"' in text and f'chart-{name}' in text and f'第 {index} 张幻灯片' in text
        for pair in ('(-6, 4)', '(-3, )', '(2, -3)', '(8, 6)', '(-7, -2)', '(-1, 4)', '(3, 2)', '(7, -6)'):
            assert pair in text
        assert image.read_bytes().startswith(b'\xff\xd8\xff')
        captures.append({'phase': phase, 'slide': index, 'axSha256': sha(ax), 'screenshotSha256': sha(image)})
result = {'scope': 'Four fixed-axis sparse scatter charts from an independently generated, PowerPoint-saved source received existing-point X/Y edits through Parser v2. PowerPoint opened the edited file, saved a copy, closed and reopened it once. All caches, worksheet coordinates and blank positions were checked; eight retained UI captures show source point labels. Screenshots were not quantitatively measured, and the captures do not independently prove provenance. Blank insertion/removal, native interactive workbook editing, other producers and further reopen rounds remain unverified.',
    'applicationVersion': plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],
    'generatorSha256': sha(ROOT / 'scripts/prepare-native-scatter-sparse-save.mjs'), 'verifierSha256': sha(Path(__file__)), 'manifestSha256': sha(BASE / 'manifest.json'), 'packages': packages, 'captures': captures}
if sys.argv[1:] == ['--record']:
    RECORD.write_text(json.dumps(result, indent=2) + '\n')
else:
    record = Path(sys.argv[1]) if len(sys.argv) == 2 else RECORD
    assert json.loads(record.read_text()) == result
print(json.dumps({'verified': True, 'charts': 4, 'coordinateEdits': 8, 'reopenRounds': 1}))
