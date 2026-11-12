#!/usr/bin/env python3
"""Check retained bytes and independent OOXML properties; UI provenance is observational."""
import hashlib
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
record = json.loads((ROOT / 'docs/testing/edit-transform-powerpoint-evidence.json').read_text())
for item in record['artifacts'].values():
    assert hashlib.sha256((ROOT / item['path']).read_bytes()).hexdigest() == item['sha256']

def read(path):
    with zipfile.ZipFile(ROOT / path) as archive:
        root = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
    shapes = {}
    for shape in root.findall('.//p:sp', NS):
        name = shape.find('p:nvSpPr/p:cNvPr', NS).get('name')
        xfrm = shape.find('p:spPr/a:xfrm', NS)
        shapes[name] = (xfrm.attrib, xfrm.find('a:off', NS).attrib,
                        xfrm.find('a:ext', NS).attrib,
                        ''.join(t.text or '' for t in shape.findall('.//a:t', NS)))
    return shapes

source = read(record['artifacts']['source']['path'])
saved = read(record['artifacts']['saved']['path'])
assert set(source) == set(saved) == {'Transform target', 'reference'}
before, after = source['Transform target'], saved['Transform target']
assert before[0] == after[0] == {'rot': '1800000', 'flipH': '1'}
assert before[1] == {'x': '1905000', 'y': '1397000'}
assert after[1] == {'x': '1927302', 'y': '1397000'}
assert before[2] == after[2] == {'cx': '3302000', 'cy': '1524000'}
assert before[3] == after[3] == 'Transform sample'
assert source['reference'] == saved['reference']
assert source['reference'][3] == 'Rotation 30 degrees; horizontal flip'
print('Retained transform evidence verified.')
