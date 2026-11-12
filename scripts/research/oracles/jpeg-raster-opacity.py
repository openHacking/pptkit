"""Verify named JPEG inputs, retained raster alpha and strict-policy results."""
import hashlib
import json
import sys
from pathlib import Path

from PIL import Image

if len(sys.argv) != 2:
    raise SystemExit('Usage: python3 jpeg-raster-opacity.py CAPTURE_DIRECTORY')
root = Path(sys.argv[1])
manifest = json.loads((root / 'evidence.json').read_text())
digest = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
assert digest(Path('scripts/research/jpeg-raster-opacity-observation.mjs')) == manifest['producerSha256']
for name, expected in manifest['fixtures'].items():
    assert digest(Path('tests/fixtures') / name) == expected
rejected = 0
for engine in manifest['observations']:
    assert engine['networkRequests'] == []
    assert len(engine['results']) == 6
    for item in engine['results']:
        if item['status'] == 'rejected':
            assert item['code'] == 'raster-resource-unavailable'
            rejected += 1
            continue
        path = root / item['png']
        assert digest(path) == item['pngSha256']
        image = Image.open(path).convert('RGBA')
        assert image.size == (6, 4)
        nonopaque = sum(pixel[3] != 255 for pixel in image.get_flattened_data())
        assert nonopaque == item['nonopaque']
        if item['policy'] == 'require-opaque':
            assert nonopaque == 0
        changed = image.copy(); changed.putpixel((0, 0), (0, 0, 0, 255 if image.getpixel((0, 0))[3] != 255 else 0))
        assert sum(pixel[3] != 255 for pixel in changed.get_flattened_data()) != nonopaque
print(json.dumps({'engines': len(manifest['observations']), 'rejected': rejected, 'verifiedRasters': sum(item['status'] == 'rendered' for engine in manifest['observations'] for item in engine['results'])}))
