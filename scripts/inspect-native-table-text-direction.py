"""Verify the retained PowerPoint observation; requires Pillow and its local artifacts."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from PIL import Image

repo = Path(__file__).resolve().parent.parent
record = json.loads((repo / 'docs/testing/table-text-direction-powerpoint-evidence.json').read_text())
root = repo / record['artifactDirectory']
for name, digest in record['sha256'].items():
    path = repo / record['fixture'] if name == 'fixture' else root / name
    assert hashlib.sha256(path.read_bytes()).hexdigest() == digest, name
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
for name, key in [('source.pptx', 'sourceDirections'), ('native-round-1.pptx', 'savedDirections')]:
    with ZipFile(root / name) as archive:
        xml = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
    cells = xml.findall('.//a:tbl/a:tr/a:tc', ns)
    assert [''.join(c.itertext()) for c in cells] == ['DEFAULT', 'CLOCKWISE', 'COUNTER', 'HORIZONTAL']
    assert [c.find('a:tcPr', ns).get('vert') for c in cells] == record['packageObservation'][key]
first, reopen = (Image.open(root / name).convert('RGB') for name in ['first.png', 'reopen.png'])
assert first.size == reopen.size == (1600, 900)
for region in record['glyphMaskComparison']['regions']:
    x0, y0, x1, y1 = region['bounds']
    masks = [{(x, y) for y in range(y0, y1) for x in range(x0, x1) if max(image.getpixel((x, y))) < 80} for image in (first, reopen)]
    assert [len(mask) for mask in masks] == [region['first'], region['reopen']], region['name']
    assert len(masks[0] ^ masks[1]) == region['symmetricDifference'], region['name']
for name in ['first.ax.txt', 'reopen.ax.txt']:
    text = (root / name).read_text()
    assert all(value in text for value in ['direction-table', 'DEFAULT', 'CLOCKWISE', 'COUNTER', 'HORIZONTAL'])
print('PowerPoint table direction evidence verified')
