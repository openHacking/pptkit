#!/usr/bin/env python3
"""Check retained PowerPoint output independently of PPTKit; never drives UI."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()

def inspect(path, headline, offset=('2286000', '1905000')):
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
    shapes = root.findall('.//p:sp', NS)
    named = {shape.find('p:nvSpPr/p:cNvPr', NS).get('name'): shape for shape in shapes}
    if set(named) != {'headline', 'shape'} or len(shapes) != 2: raise ValueError('shape inventory differs')
    text = lambda shape: ''.join(node.text or '' for node in shape.findall('.//a:t', NS))
    if text(named['headline']) != headline or text(named['shape']) != 'Edited shape': raise ValueError('text differs')
    groups = root.findall('.//p:grpSp', NS)
    if len(groups) != 1 or groups[0].find('p:nvGrpSpPr/p:cNvPr', NS).get('name') != 'group' or groups[0].find('p:sp/p:nvSpPr/p:cNvPr', NS).get('name') != 'shape': raise ValueError('group ownership differs')
    xfrm = groups[0].find('p:grpSpPr/a:xfrm', NS)
    expected = {'off': {'x': offset[0], 'y': offset[1]}, 'ext': {'cx': '2540000', 'cy': '1143000'}, 'chOff': {'x': '0', 'y': '0'}, 'chExt': {'cx': '2540000', 'cy': '1143000'}}
    if any(xfrm.find('a:' + key, NS).attrib != value for key, value in expected.items()): raise ValueError('group geometry differs')
    return {name: ET.tostring(shape.find('p:spPr/a:xfrm', NS)) for name, shape in named.items()}

def verify(local=False):
    record = json.loads((ROOT / 'docs/testing/edit-operation-powerpoint-evidence.json').read_text())
    for item in record['artifacts'].values():
        if sha(ROOT / item['path']) != item['sha256']: raise ValueError('artifact hash differs')
    inspect(ROOT / record['artifacts']['source']['path'], 'Edited headline')
    inspect(ROOT / record['artifacts']['output']['path'], 'Edited native')
    if local:
        directory = ROOT / record['artifactDirectory']
        for name, value in record['captures'].items():
            if sha(directory / name) != value: raise ValueError('capture hash differs')
        for name, headline in [('first.ax.txt', 'Edited headline'), ('reopen.ax.txt', 'Edited native')]:
            state = (directory / name).read_text()
            if any(value not in state for value in ('native-edit', headline, 'Edited shape', 'group')): raise ValueError('captured state differs')

if __name__ == '__main__':
    if sys.argv[1:] not in ([], ['--local']): raise SystemExit('usage: verify-edit-operation-native.py [--local]')
    verify(bool(sys.argv[1:]))
    print('verified native edit-operation sample' + (' and local captures' if sys.argv[1:] else ''))
