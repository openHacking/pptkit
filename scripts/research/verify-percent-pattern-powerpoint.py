#!/usr/bin/env python3
"""Check the retained, single PowerPoint save/reopen observation; no visual-parity claim."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
ARTIFACTS = ROOT / '.tmp/native-percent-patterns'
RECORD = ROOT / 'docs/testing/percent-pattern-powerpoint-evidence.json'
VALUES = (5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90)
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
FILES = ('percent-patterns.pptx', 'native-round-1.pptx', 'first.ax.txt', 'reopen.ax.txt', 'first.jpg', 'reopen.jpg')

def check():
    hashes = {name: hashlib.sha256((ARTIFACTS / name).read_bytes()).hexdigest() for name in FILES}
    fixture = ROOT / 'tests/fixtures/percent-pattern-powerpoint-resave.pptx'
    assert hashlib.sha256(fixture.read_bytes()).hexdigest() == hashes['native-round-1.pptx']
    for filename in FILES[:2]:
        with zipfile.ZipFile(ARTIFACTS / filename) as package:
            root = ET.fromstring(package.read('ppt/slides/slide1.xml'))
        shapes = root.findall('.//p:sp', NS)
        assert len(shapes) == len(VALUES)
        for shape, value in zip(shapes, VALUES):
            assert shape.find('p:nvSpPr/p:cNvPr', NS).get('name') == f'percent{value}'
            fill = shape.find('p:spPr/a:pattFill', NS)
            assert fill is not None and fill.get('prst') == f'pct{value}'
            assert fill.find('a:fgClr/a:srgbClr', NS).get('val') == '000000'
            assert fill.find('a:bgClr/a:srgbClr', NS).get('val') == 'FFFFFF'
    for filename, deck in [('first.ax.txt', 'percent-patterns.pptx'), ('reopen.ax.txt', 'native-round-1.pptx')]:
        state = (ARTIFACTS / filename).read_text()
        assert f'/native-percent-patterns/{deck}' in state
        assert all(f'形状 (disabled) percent{value}\n' in state for value in VALUES)
    for filename in FILES[-2:]:
        assert (ARTIFACTS / filename).read_bytes().startswith(b'\xff\xd8\xff')
    return {'scope': 'One authored slide, opaque black/white rectangular shapes; one PowerPoint save and reopen. UI captures and XML are retained; visual parity, color/alpha variations, interactive edits and other producers are not established.', 'presetValues': list(VALUES), 'reopenRounds': 1, 'artifactDirectory': '.tmp/native-percent-patterns', 'sha256': hashes}

observed = check()
if sys.argv[1:] == ['--record']:
    RECORD.write_text(json.dumps(observed, indent=2, ensure_ascii=False) + '\n')
else:
    assert json.loads(RECORD.read_text()) == observed
print(json.dumps({'verified': True, 'presets': len(VALUES), 'reopenRounds': 1}))
