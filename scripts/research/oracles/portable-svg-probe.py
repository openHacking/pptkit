"""Check one controlled PPTKit SVG in Chromium and independent MuPDF.

The assertions intentionally describe this sample, not general SVG support.
"""
import hashlib
import json
import sys
from pathlib import Path
from xml.etree import ElementTree

import fitz
from PIL import Image

root = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/portable-svg-probe')
svg = root / 'specimen.svg'
doc = fitz.open(svg)
try:
    pixmap = doc[0].get_pixmap(alpha=False)
    pixmap.save(root / 'mupdf.png')
finally:
    doc.close()
browser = Image.open(root / 'chromium.png').convert('RGB')
independent = Image.open(root / 'mupdf.png').convert('RGB')


def dark(im, box):
    return sum(max(pixel) < 150 for pixel in im.crop(box).get_flattened_data())


def close(actual, expected, tolerance):
    return all(abs(a - b) <= tolerance for a, b in zip(actual, expected))


def check(a, b):
    assert a.size == b.size == (400, 260)
    regions = {'native': (10, 10, 190, 40), 'rich': (10, 50, 190, 80), 'table': (10, 95, 190, 140)}
    counts = {name: [dark(a, box), dark(b, box)] for name, box in regions.items()}
    assert min(counts['native']) > 100
    assert counts['rich'][0] > 100 and counts['rich'][1] == 0
    assert counts['table'][0] > 100 and counts['table'][1] == 0
    for point, expected in [((225, 165), (255, 0, 0)), ((280, 180), (127, 0, 128)), ((350, 180), (127, 127, 255))]:
        assert close(a.getpixel(point), expected, 2) and close(b.getpixel(point), expected, 2)
    for point in [(240, 50), (315, 50)]:
        assert close(a.getpixel(point), b.getpixel(point), 8)
        assert not close(a.getpixel(point), (255, 255, 255), 20)
    assert close(a.getpixel((30, 115)), (255, 255, 0), 1)
    assert close(b.getpixel((30, 115)), (255, 255, 255), 1)
    return counts


counts = check(browser, independent)
image = ElementTree.parse(svg).find('.//{http://www.w3.org/2000/svg}image')
assert image is not None and image.get('href', '').startswith('data:image/png;base64,')
for key, expected in [('x', 220), ('y', -110), ('width', 120), ('height', 360)]:
    assert abs(float(image.get(key)) - expected) < 1e-9
for index in range(5):
    altered = browser.copy()
    altered.paste((255, 255, 255), [(10, 10, 190, 40), (10, 50, 190, 80), (10, 95, 190, 140), (220, 10, 340, 130), (220, 160, 370, 230)][index])
    try:
        check(altered, independent)
    except AssertionError:
        continue
    raise AssertionError(f'negative control {index} was accepted')

assert json.loads((root / 'requests.json').read_text()) == []
warnings = json.loads((root / 'warnings.json').read_text())
assert {(w['code'], w['elementId']) for w in warnings if w['code'].startswith('svg-')} == {
    ('svg-system-font-dependency', 'native'), ('svg-system-font-dependency', 'rich'),
    ('svg-foreign-object-dependency', 'rich'), ('svg-system-font-dependency', 'table'),
    ('svg-foreign-object-dependency', 'table'),
}
files = ['specimen.svg', 'warnings.json', 'chromium.png', 'mupdf.png', 'requests.json', 'browser-version.txt']
result = {'chromium': (root / 'browser-version.txt').read_text().strip(), 'pymupdf': fitz.VersionBind,
          'mupdf': fitz.VersionFitz, 'darkPixels': counts, 'negativeControlsRejected': 5,
          'sha256': {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in files}}
print(json.dumps(result, indent=2))
