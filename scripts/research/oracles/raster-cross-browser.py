"""Independently inspect retained cross-engine PNGs and reject altered controls."""
import hashlib
import json
import sys
from pathlib import Path

from PIL import Image

if len(sys.argv) != 2:
    raise SystemExit('Usage: python3 raster-cross-browser.py CAPTURE_DIRECTORY')
root = Path(sys.argv[1])
manifest = json.loads((root / 'evidence.json').read_text())
font = Path('packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2')
assert hashlib.sha256(font.read_bytes()).hexdigest() == manifest['fontSha256']
producer = Path('scripts/research/raster-cross-browser-observation.mjs')
assert hashlib.sha256(producer.read_bytes()).hexdigest() == manifest['producerSha256']


def inspect(plain, covered):
    assert plain.size == covered.size == (340, 120)
    a, b = list(plain.get_flattened_data()), list(covered.get_flattened_data())
    changed = sum(x != y for x, y in zip(a, b))
    dark = sum(r < 120 and g < 120 and bl < 120 and alpha > 0 for r, g, bl, alpha in b)
    overlay, transparent = covered.getpixel((40, 90)), covered.getpixel((320, 110))
    assert changed > 100 and dark > 100
    assert abs(overlay[0] - 127) <= 1 and overlay[1:] == (0, 128, 255)
    assert transparent == (0, 0, 0, 0)
    return changed, dark, overlay, transparent


for entry in manifest['observations']:
    name = entry['engine']
    for suffix in ('plain', 'covered'):
        path = root / f'{name}-{suffix}.png'
        assert hashlib.sha256(path.read_bytes()).hexdigest() == entry[f'{suffix}Sha256']
    plain = Image.open(root / f'{name}-plain.png').convert('RGBA')
    covered = Image.open(root / f'{name}-covered.png').convert('RGBA')
    changed, dark, overlay, transparent = inspect(plain, covered)
    assert changed == entry['changedPixels'] and dark == entry['darkPixels']
    assert list(overlay) == entry['overlay'] and list(transparent) == entry['transparent']
    assert entry['missingFont'] == 'raster-font-unavailable' and entry['networkRequests'] == []
    controls = [plain.copy(), Image.new('RGBA', covered.size), covered.copy()]
    controls[2].putpixel((40, 90), (0, 0, 0, 0))
    for altered in controls:
        try:
            inspect(plain, altered)
        except AssertionError:
            continue
        raise AssertionError(f'{name}: negative control was accepted')
print(json.dumps({'engines': len(manifest['observations']), 'negativeControlsRejected': 3 * len(manifest['observations'])}))
