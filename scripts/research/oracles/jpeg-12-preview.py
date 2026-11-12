"""Check one Chromium SVG observation against the retained PowerPoint screenshot."""
from pathlib import Path
import base64
import hashlib
import json
import xml.etree.ElementTree as ET
from PIL import Image

root = Path(__file__).resolve().parents[3]
work = root / '.tmp/native-jpeg-12'
report = json.loads((work / 'crop-preview-report.json').read_text())
digest = lambda data: hashlib.sha256(data).hexdigest()
assert report['inputSha256'] == digest((work / 'crop.pptx').read_bytes())
assert report['svgSha256'] == digest((work / 'crop-preview.svg').read_bytes())
assert report['screenshotSha256'] == digest((work / 'crop-preview.png').read_bytes())
assert report['requests'] == [] and len(report['decoded']) == 3
assert all(entry == {'ok': True, 'width': 3, 'height': 2, 'firstPixel': [0, 0, 0, 0]} for entry in report['decoded'])
assert report['control']['firstPixel'][3] == 255
assert [item['code'] for item in report['warnings']] == ['jpeg-preview-precision-unverified'] * 2
svg = ET.fromstring((work / 'crop-preview.svg').read_bytes())
images = svg.findall('.//{http://www.w3.org/2000/svg}image')
assert len(images) == 3
expected = ['jpeg-12bit-sequential.jpg', 'jpeg-12bit-progressive.jpg', 'jpeg-12bit-progressive.jpg']
for node, name in zip(images, expected):
    url = node.attrib['href']; assert url.startswith('data:image/jpeg;base64,')
    assert base64.b64decode(url.split(',', 1)[1]) == (root / 'tests/fixtures' / name).read_bytes()
browser = Image.open(work / 'crop-preview.png').convert('RGB')
native = Image.open(work / 'crop-reopen.png').convert('RGB')
assert browser.size == native.size == (1920, 1080)
for y in (320, 520, 720):
    for x in (110, 710, 1430):
        assert browser.getpixel((x, y)) == (255, 255, 255)
        assert native.getpixel((x, y)) != (255, 255, 255)
for name in ('crop.pptx', 'crop-preview.svg', 'crop-preview.png', 'crop-preview-report.json', 'crop-reopen.png'):
    print(digest((work / name).read_bytes()), name)
