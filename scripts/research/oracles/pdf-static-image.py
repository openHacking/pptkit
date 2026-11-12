"""Independently inspect static and deliberately misattested PDF captures."""
import hashlib
import base64
import json
import sys
from pathlib import Path
import fitz
from PIL import Image

if len(sys.argv) != 2:
    raise SystemExit('Usage: python3 pdf-static-image.py CAPTURE_DIRECTORY')
root = Path(sys.argv[1])
manifest = json.loads((root / 'evidence.json').read_text())
digest = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
assert digest(Path('scripts/research/pdf-static-image-observation.mjs')) == manifest['producerSha256']
assert digest(Path('packages/pdf-renderer/src/rendering/print.ts')) == manifest['adapterSourceSha256']
assert digest(Path('packages/pdf-renderer/src/rendering/image-data.ts')) == manifest['imageDataSha256']
assert digest(Path('packages/media-inspection/src/jpeg/structure.ts')) == manifest['jpegSourceSha256']
for name, expected in manifest['mediaSources'].items():
    assert digest(Path('packages/media-inspection/src/images') / name) == expected
for name, value in manifest['fixtures'].items():
    assert digest(Path('tests/fixtures') / name) == value
for name, frames in [('transparent-replacement.png', 1), ('two-frame-apng.png', 2)]:
    with Image.open(Path('tests/fixtures') / name) as image:
        assert image.n_frames == frames
assert manifest['rejected'] == 'pdf-image-not-static'
assert manifest['automaticRejected'] == 'pdf-image-not-static'
assert manifest['gifRejected'] == 'pdf-image-not-static'
for name, frames in [('still', 1), ('animated', 2)]:
    path = root / (name + '.gif')
    assert digest(path) == manifest['gifFixtures'][name]
    with Image.open(path) as image:
        assert image.n_frames == frames
assert [(item['slideId'], item['slideIndex'], item['mimeType']) for item in manifest['calls']] == [('s1', 1, 'image/png'), ('s0', 0, 'image/png')]
data_url = b'data:image/png;base64,' + base64.b64encode((Path('tests/fixtures') / 'transparent-replacement.png').read_bytes())
assert all(item['dataSha256'] == hashlib.sha256(data_url).hexdigest() for item in manifest['calls'])
for name in ('still', 'misattested', 'automatic', 'gif'):
    item = manifest[name]
    assert digest(root / (name + '.pdf')) == item['sha256']
    codes = [warning['code'] for warning in item['warnings']]
    assert 'pdf-image-frame-unverified' in codes
    assert ('pdf-image-static-structure-checked' if name in ('automatic', 'gif') else 'pdf-image-static-attested') in codes
def check_still(document):
    assert len(document) == 2
    for page in document:
        pixels = page.get_pixmap(alpha=False)
        assert pixels.pixel(5, 10) == (255, 0, 0) and pixels.pixel(35, 10) == (255, 255, 255)
with fitz.open(root / 'still.pdf') as document:
    check_still(document)
with fitz.open(root / 'automatic.pdf') as document:
    check_still(document)
with fitz.open(root / 'still.pdf') as altered:
    altered[0].draw_rect(fitz.Rect(0, 0, 20, 20), color=(1, 1, 1), fill=(1, 1, 1), overlay=True)
    changed = altered.tobytes()
with fitz.open(stream=changed, filetype='pdf') as altered:
    try:
        check_still(altered)
    except AssertionError:
        pass
    else:
        raise AssertionError('Covered static image unexpectedly passed')
with fitz.open(root / 'misattested.pdf') as document:
    assert len(document) == 1
with fitz.open(root / 'gif.pdf') as document:
    assert len(document) == 1
print(json.dumps({'mupdf': fitz.VersionBind, 'staticPages': 2, 'gifFrames': [1, 2], 'apngRejectedByCaller': True, 'falseAttestationAccepted': True, 'automaticApngRejected': True, 'automaticGifRejected': True, 'negativeControls': 1}))
