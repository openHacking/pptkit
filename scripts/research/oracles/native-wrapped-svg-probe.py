"""Record MuPDF visibility and extraction of wrapped mixed SVG text."""
import hashlib
import json
import sys
from pathlib import Path

import fitz
from PIL import Image

root = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/native-wrapped-svg-probe')
results = {}
for name in ('browser', 'native'):
    doc = fitz.open(root / f'{name}.svg')
    try:
        page = doc[0]
        text = page.get_text()
        lines = [line for block in page.get_text('dict')['blocks'] if 'lines' in block for line in block['lines']]
        page.get_pixmap(alpha=False).save(root / f'{name}-mupdf.png')
    finally:
        doc.close()
    image = Image.open(root / f'{name}-mupdf.png').convert('RGB')
    red = sum(r > 150 and g < 120 and b < 120 for r, g, b in image.get_flattened_data())
    blue = sum(b > 150 and r < 120 and g < 120 for r, g, b in image.get_flattened_data())
    results[name] = {'text': text, 'lineCount': len(lines), 'redPixels': red, 'bluePixels': blue}
assert results['browser']['redPixels'] == results['browser']['bluePixels'] == 0
assert results['native']['redPixels'] > 25 and results['native']['bluePixels'] > 25
assert results['native']['lineCount'] >= 2
assert ''.join(results['native']['text'].split()) == 'AlphaBetaGamma'
assert not results['browser']['text'].strip()
warnings = [w['code'] for w in json.loads((root / 'native-warnings.json').read_text())]
assert 'svg-native-mixed-text-approximation' in warnings
files = [f'{name}{suffix}' for name in ('browser', 'native') for suffix in ('.svg', '-warnings.json', '-mupdf.png')]
print(json.dumps({'pymupdf': fitz.VersionBind, 'mupdf': fitz.VersionFitz, 'observed': results,
                  'sha256': {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in files}}, indent=2))
