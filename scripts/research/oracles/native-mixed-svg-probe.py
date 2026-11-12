"""Compare actual MuPDF output for default and opt-in mixed-run SVGs."""
import hashlib
import json
import sys
from pathlib import Path

import fitz
from PIL import Image

root = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/native-mixed-svg-probe')
results = {}
for name in ('browser', 'native'):
    d = fitz.open(root / f'{name}.svg')
    try:
        page = d[0]
        text = page.get_text()
        page.get_pixmap(alpha=False).save(root / f'{name}-mupdf.png')
    finally:
        d.close()
    im = Image.open(root / f'{name}-mupdf.png').convert('RGB')
    red = sum(r > 150 and g < 120 and b < 120 for r, g, b in im.get_flattened_data())
    blue = sum(b > 150 and r < 120 and g < 120 for r, g, b in im.get_flattened_data())
    results[name] = {'text': text, 'redPixels': red, 'bluePixels': blue}
assert results['browser']['redPixels'] == results['browser']['bluePixels'] == 0
assert results['native']['redPixels'] > 25 and results['native']['bluePixels'] > 25
assert 'Alpha' in results['native']['text'] and 'Beta' in results['native']['text']
assert 'Alpha' not in results['browser']['text']
for name, expected in [('browser', 'svg-foreign-object-dependency'), ('native', 'svg-native-mixed-text-approximation')]:
    assert expected in [w['code'] for w in json.loads((root / f'{name}-warnings.json').read_text())]
files = [f'{name}{suffix}' for name in ('browser', 'native') for suffix in ('.svg', '-warnings.json', '-mupdf.png')]
print(json.dumps({'pymupdf': fitz.VersionBind, 'mupdf': fitz.VersionFitz, 'observed': results,
                  'sha256': {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in files}}, indent=2))
