"""Inspect named Chromium PDF JPEG streams and MuPDF rendering independently."""
import hashlib
import json
import sys
from pathlib import Path
import fitz

if len(sys.argv) != 2:
    raise SystemExit('Usage: python3 pdf-jpeg-opacity.py CAPTURE_DIRECTORY')
root = Path(sys.argv[1])
manifest = json.loads((root / 'evidence.json').read_text())
digest = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
assert digest(Path('scripts/research/pdf-jpeg-opacity-observation.mjs')) == manifest['producerSha256']
assert digest(Path('packages/pdf-renderer/src/rendering/print.ts')) == manifest['adapterSourceSha256']
assert digest(Path('packages/pdf-renderer/src/rendering/jpeg-prefix.ts')) == manifest['jpegProbeSha256']
assert digest(Path('packages/pdf-renderer/src/rendering/image-data.ts')) == manifest['imageDataSha256']
assert digest(Path('packages/media-inspection/src/jpeg/structure.ts')) == manifest['structureSourceSha256']
for name, expected in manifest['fixtures'].items():
    assert digest(Path('tests/fixtures') / name) == expected
assert len(manifest['results']) == 12
assert {(item['name'], item['policy']) for item in manifest['results']} == {(name, policy) for name in manifest['fixtures'] for policy in ('browser', 'require-opaque', 'require-8-bit', 'require-valid-8-bit')}
assert digest(root / 'duplicate-frame.jpg') == manifest['duplicate']['sha256']
assert manifest['duplicate']['result'] == 'pdf-jpeg-structure-unavailable'

def precision(data):
    assert data[:2] == b'\xff\xd8'
    offset = 2
    while offset < len(data):
        assert data[offset] == 0xff
        while data[offset] == 0xff:
            offset += 1
        marker = data[offset]; offset += 1
        assert marker not in (0xd9, 0xda) and offset + 2 <= len(data)
        length = int.from_bytes(data[offset:offset + 2], 'big')
        assert length >= 2 and offset + length <= len(data)
        if marker in (0xc0, 0xc1, 0xc2):
            return data[offset + 2]
        offset += length
    raise AssertionError('Missing JPEG frame header')

fitz.TOOLS.mupdf_display_errors(False)
observed = {}
def check_12bit(document, expected_hash):
    assert len(document) == 1
    images = document[0].get_images(full=True)
    assert len(images) == 1 and images[0][8] == 'DCTDecode'
    xref = images[0][0]
    assert document.xref_get_key(xref, 'BitsPerComponent') == ('int', '8')
    assert hashlib.sha256(document.xref_stream_raw(xref)).hexdigest() == expected_hash
    fitz.TOOLS.mupdf_warnings()
    pixels = document[0].get_pixmap(matrix=fitz.Matrix(4, 4), alpha=False)
    assert pixels.pixel(10, 8) == (0, 0, 0)
    assert 'Unsupported JPEG data precision 12' in fitz.TOOLS.mupdf_warnings()

for item in manifest['results']:
    source = (Path('tests/fixtures') / item['name']).read_bytes()
    bits = precision(source)
    assert bits == (8 if item['name'] == 'exif-orientation6.jpg' else 12)
    if item['status'] == 'rejected':
        assert bits == 12
        assert (item['policy'], item['code']) in [('require-opaque', 'pdf-resource-unavailable'), ('require-8-bit', 'pdf-jpeg-unsupported-precision'), ('require-valid-8-bit', 'pdf-jpeg-unsupported-precision')]
        continue
    path = root / item['file']
    assert digest(path) == item['sha256']
    with fitz.open(path) as document:
        assert len(document) == 1
        if bits == 12:
            assert item['policy'] == 'browser'
            check_12bit(document, manifest['fixtures'][item['name']])
        else:
            pix = document[0].get_pixmap(matrix=fitz.Matrix(4, 4), alpha=False)
            assert pix.pixel(10, 8) not in ((0, 0, 0), (255, 255, 255))
    observed[f"{item['name']}:{item['policy']}"] = item['status']

path = root / 'jpeg-12bit-sequential.jpg-browser.pdf'
with fitz.open(path) as changed:
    xref = changed[0].get_images()[0][0]
    changed.xref_set_key(xref, 'BitsPerComponent', '12')
    altered = changed.tobytes()
with fitz.open(stream=altered, filetype='pdf') as changed:
    try:
        check_12bit(changed, manifest['fixtures']['jpeg-12bit-sequential.jpg'])
    except AssertionError:
        pass
    else:
        raise AssertionError('Changed PDF dictionary unexpectedly passed')
duplicate = (root / 'duplicate-frame.jpg').read_bytes()
assert precision(duplicate) == 8
offset, frames = 2, []
while offset < len(duplicate):
    assert duplicate[offset] == 0xff
    while duplicate[offset] == 0xff:
        offset += 1
    marker = duplicate[offset]; offset += 1
    if marker == 0xda:
        break
    length = int.from_bytes(duplicate[offset:offset + 2], 'big')
    assert length >= 2 and offset + length <= len(duplicate)
    if marker in (0xc0, 0xc1, 0xc2):
        frames.append(duplicate[offset - 2:offset + length])
    offset += length
assert len(frames) == 2 and frames[0] == frames[1]
print(json.dumps({'mupdf': fitz.VersionBind, 'rendered': observed, 'strictRejections': 6, 'duplicateFrameRejected': True, 'negativeControls': 1}))
