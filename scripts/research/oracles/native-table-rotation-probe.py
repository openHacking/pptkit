"""Inspect actual MuPDF text directions and colored pixels for a table."""
import hashlib
import json
import math
import sys
from pathlib import Path

import fitz
from PIL import Image

root = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/native-table-rotation-probe')
results = {}
for name in ('browser', 'native'):
    doc = fitz.open(root / f'{name}.svg')
    try:
        page = doc[0]
        lines = [line for block in page.get_text('dict')['blocks'] if 'lines' in block for line in block['lines']]
        observed = [{'text': ''.join(span['text'] for span in line['spans']), 'angle': math.degrees(math.atan2(line['dir'][1], line['dir'][0]))} for line in lines]
        page.get_pixmap(alpha=False).save(root / f'{name}-mupdf.png')
    finally:
        doc.close()
    im = Image.open(root / f'{name}-mupdf.png').convert('RGB')
    pixels = im.load()
    red = [(x, y) for y in range(im.height) for x in range(im.width) if (lambda c: c[0] > 150 and c[1] < 120 and c[2] < 120)(pixels[x, y])]
    blue = [(x, y) for y in range(im.height) for x in range(im.width) if (lambda c: c[2] > 150 and c[0] < 120 and c[1] < 120)(pixels[x, y])]
    bounds = lambda points: None if not points else [min(x for x, _ in points), min(y for _, y in points), max(x for x, _ in points), max(y for _, y in points)]
    results[name] = {'lines': observed, 'redPixels': len(red), 'bluePixels': len(blue), 'redBounds': bounds(red), 'blueBounds': bounds(blue)}
assert results['browser']['redPixels'] == results['browser']['bluePixels'] == 0
assert results['native']['redPixels'] > 25 and results['native']['bluePixels'] > 25
assert 20 <= results['native']['redBounds'][0] <= results['native']['redBounds'][2] < 120
assert 120 <= results['native']['blueBounds'][0] <= results['native']['blueBounds'][2] < 220
for color in ('red', 'blue'):
    assert 20 <= results['native'][color + 'Bounds'][1] <= results['native'][color + 'Bounds'][3] < 180
lines = {line['text']: line['angle'] for line in results['native']['lines']}
assert abs(lines['Clockwise'] - 90) < 0.01 and abs(lines['Counter'] + 90) < 0.01
warnings = [w['code'] for w in json.loads((root / 'native-warnings.json').read_text())]
assert 'table-text-direction-approximation' in warnings and 'svg-native-table-approximation' in warnings
assert 'svg-native-table-fallback' not in warnings
files = [f'{name}{suffix}' for name in ('browser', 'native') for suffix in ('.svg', '-warnings.json', '-mupdf.png')]
print(json.dumps({'pymupdf': fitz.VersionBind, 'mupdf': fitz.VersionFitz, 'observed': results,
                  'sha256': {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in files}}, indent=2))
