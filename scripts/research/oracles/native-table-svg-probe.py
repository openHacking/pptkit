"""Observe a merged native SVG table through the independent MuPDF engine."""
import hashlib
import json
import sys
from pathlib import Path
from xml.etree import ElementTree

import fitz
from PIL import Image

root = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/native-table-svg-probe')
results = {}
for name in ('browser', 'native'):
    doc = fitz.open(root / f'{name}.svg')
    try:
        text = doc[0].get_text()
        doc[0].get_pixmap(alpha=False).save(root / f'{name}-mupdf.png')
    finally:
        doc.close()
    image = Image.open(root / f'{name}-mupdf.png').convert('RGB')
    points = [(30, 25), (30, 65), (205, 95)]
    results[name] = {'text': text, 'samples': [image.getpixel(point) for point in points]}
assert results['browser']['samples'] == [(255, 255, 255)] * 3
assert results['native']['samples'] == [(255, 255, 0), (0, 255, 255), (255, 176, 176)]
assert all(word in results['native']['text'] for word in ('Header', 'Left', 'Right'))
assert not results['browser']['text'].strip()
tree = ElementTree.parse(root / 'native.svg')
ns = {'s': 'http://www.w3.org/2000/svg'}
cells = [element for element in tree.findall('.//s:g', ns) if element.get('data-pptkit-cell')]
assert len(cells) == 3 and cells[0].get('aria-colspan') == '2'
assert all(element.get('clip-path', '').startswith('url(#') for element in cells)
for name, code in [('browser', 'svg-foreign-object-dependency'), ('native', 'svg-native-table-approximation')]:
    assert code in [item['code'] for item in json.loads((root / f'{name}-warnings.json').read_text())]
files = [f'{name}{suffix}' for name in ('browser', 'native') for suffix in ('.svg', '-warnings.json', '-mupdf.png')]
print(json.dumps({'pymupdf': fitz.VersionBind, 'mupdf': fitz.VersionFitz, 'observed': results,
                  'sha256': {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in files}}, indent=2))
