"""Record MuPDF's observed face and raster with and without one embedded WOFF2."""
import hashlib
import json
import sys
from pathlib import Path

import fitz

root = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/embedded-font-mupdf-probe')
observed = {}
for name in ('plain', 'embedded'):
    doc = fitz.open(root / f'{name}.svg')
    try:
        page = doc[0]
        spans = [span for block in page.get_text('dict')['blocks'] if 'lines' in block for line in block['lines'] for span in line['spans']]
        pixmap = page.get_pixmap(alpha=False)
        pixmap.save(root / f'{name}-mupdf.png')
        observed[name] = {'spans': [{'font': span['font'], 'text': span['text']} for span in spans], 'pixelsSha256': hashlib.sha256(pixmap.samples).hexdigest()}
    finally:
        doc.close()
assert observed['plain'] == observed['embedded']
assert observed['embedded']['spans'] == [{'font': 'Times-Roman', 'text': 'iiii WWWW 12345'}]
assert 'svg-embedded-font-unverified' in [w['code'] for w in json.loads((root / 'embedded-warnings.json').read_text())]
font = Path('packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2')
files = [f'{name}{suffix}' for name in ('plain', 'embedded') for suffix in ('.svg', '-warnings.json', '-mupdf.png')]
print(json.dumps({'pymupdf': fitz.VersionBind, 'mupdf': fitz.VersionFitz, 'observed': observed,
                  'fontSha256': hashlib.sha256(font.read_bytes()).hexdigest(),
                  'sha256': {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in files}}, indent=2))
