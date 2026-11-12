#!/usr/bin/env python3
"""Independently check PowerPoint's saved run-RTL control or edited copy."""
import sys
from zipfile import ZipFile
from xml.etree import ElementTree as ET

path, mode = sys.argv[1:]
assert mode in {'unedited', 'edited', 'both-edited'}
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
with ZipFile(path) as archive:
    root = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
shapes = {}
for shape in root.findall('.//p:sp', ns):
    name = shape.find('./p:nvSpPr/p:cNvPr', ns).get('name')
    assert name not in shapes
    shapes[name] = shape
assert set(shapes) == {'ltr-control', 'ltr-run-rtl', 'rtl-control', 'rtl-run-rtl'}
for name, shape in shapes.items():
    runs = shape.findall('.//a:r', ns)
    pieces = [(run.findtext('a:t', default='', namespaces=ns), run.find('a:rPr/a:rtl', ns)) for run in runs]
    expected = ('ABC مرحبا - 456 XYZ' if mode in ('edited', 'both-edited') and name == 'ltr-run-rtl' else 'ABC مرحبا - 789 XYZ' if mode == 'both-edited' and name == 'rtl-run-rtl' else 'ABC مرحبا - 123 XYZ')
    assert ''.join(text for text, _ in pieces) == expected, (name, pieces)
    rtl_text = ''.join(text for text, rtl in pieces if rtl is not None and rtl.get('val', '1') in ('1', 'true'))
    if name.endswith('run-rtl'):
        assert 'مرحبا' in rtl_text
        assert ('456' if mode in ('edited', 'both-edited') and name == 'ltr-run-rtl' else '789' if mode == 'both-edited' and name == 'rtl-run-rtl' else '123') in rtl_text
    else:
        assert not rtl_text
print(f'{mode}: four logical texts and RTL coverage verified')
