#!/usr/bin/env python3
"""Check which geometry and identity changed when PowerPoint reordered this fixture."""
import hashlib
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
RED, BLUE = 'Red rectangle', 'Blue rectangle'


def require(condition, message):
    if not condition:
        raise ValueError(message)


def inventory(path):
    with ZipFile(path) as package:
        require(package.testzip() is None, 'PPTX CRC mismatch.')
        root = ET.fromstring(package.read('ppt/slides/slide1.xml'))
    tree = root.find(f'{P}cSld/{P}spTree')
    require([child.tag for child in tree] == [P + name for name in ('nvGrpSpPr', 'grpSpPr', 'sp', 'sp')],
            'Unexpected slide shape tree.')
    require(not any(node.tag.endswith('}extLst') for node in root.iter()), 'Unexpected slide extensions.')
    shapes = []
    for shape in tree.findall(f'{P}sp'):
        props = shape.find(f'{P}nvSpPr/{P}cNvPr')
        geometry = shape.find(f'{P}spPr')
        xfrm = geometry.find(f'{A}xfrm')
        require(not xfrm.attrib and geometry.find(f'{A}prstGeom').get('prst') == 'rect', 'Not an untransformed rectangle.')
        origin = xfrm.find(f'{A}off')
        extent = xfrm.find(f'{A}ext')
        shapes.append({
            'name': props.get('name'), 'id': props.get('id'), 'description': props.get('descr'),
            'boxEmu': [int(origin.get(axis)) for axis in ('x', 'y')] + [int(extent.get(axis)) for axis in ('cx', 'cy')],
            'fill': geometry.find(f'{A}solidFill/{A}srgbClr').get('val'),
        })
    return {'sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'shapeTree': shapes}


def verify(source, saved):
    before, after = inventory(source), inventory(saved)
    require([item['name'] for item in before['shapeTree']] == [RED, BLUE], 'Incorrect source shape order.')
    require(after['shapeTree'] == before['shapeTree'][::-1], 'Resave did not solely reverse the inventoried objects.')
    a, b = [item['boxEmu'] for item in before['shapeTree']]
    require(all(max(a[i], b[i]) < min(a[i] + a[i + 2], b[i] + b[i + 2]) for i in (0, 1)),
            'The objects do not overlap.')
    return {'source': before, 'saved': after}


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.exit('Usage: verify-powerpoint-reading-order.py SOURCE.pptx NATIVE-SAVED.pptx')
    print(json.dumps(verify(Path(sys.argv[1]), Path(sys.argv[2])), indent=2))
