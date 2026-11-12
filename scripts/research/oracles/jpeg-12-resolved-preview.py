"""Check caller-supplied 12-to-8-bit resolver assets and sampled preview."""
from pathlib import Path
import base64
import hashlib
import json
import xml.etree.ElementTree as ET
from PIL import Image

root = Path(__file__).resolve().parents[3]
work = root / '.tmp/native-jpeg-12'
digest = lambda data: hashlib.sha256(data).hexdigest()
report = json.loads((work / 'resolved-preview-report.json').read_text())
assert report['inputSha256'] == digest((work / 'crop.pptx').read_bytes())
assert report['svgSha256'] == digest((work / 'resolved-preview.svg').read_bytes())
assert report['screenshotSha256'] == digest((work / 'resolved-preview.png').read_bytes())
assert report['requests'] == [] and report['warnings'] == [] and len(report['resolved']) == 2
assert len(report['decoded']) == 3 and all(item['firstPixel'][3] == 255 for item in report['decoded'])
for name, ppm in [('sequential', 'decoded-sequential.ppm'), ('progressive', 'decoded12.ppm')]:
    raw = (work / ppm).read_bytes(); header = b'P6\n3 2\n4095\n'
    assert raw.startswith(header) and len(raw) == len(header) + 36
    samples = [int.from_bytes(raw[index:index + 2], 'big') for index in range(len(header), len(raw), 2)]
    expected = [round(value * 255 / 4095) for value in samples]
    actual = list(Image.open(work / f'decoded-{name}.png').convert('RGB').tobytes())
    assert actual == expected
svg = ET.fromstring((work / 'resolved-preview.svg').read_bytes())
images = svg.findall('.//{http://www.w3.org/2000/svg}image')
assert len(images) == 3
for node, name in zip(images, ('sequential', 'progressive', 'progressive')):
    url = node.attrib['href']; assert url.startswith('data:image/png;base64,')
    assert base64.b64decode(url.split(',', 1)[1]) == (work / f'decoded-{name}.png').read_bytes()
native = Image.open(work / 'crop-reopen.png').convert('RGB')
preview = Image.open(work / 'resolved-preview.png').convert('RGB')
def check(candidate):
    assert candidate.size == native.size == (1920, 1080)
    diffs = []
    for y in (320, 520, 720):
        for x in (110, 310, 510, 710, 910, 1110, 1430, 1550, 1700):
            diffs.append(max(abs(a - b) for a, b in zip(native.getpixel((x, y)), candidate.getpixel((x, y)))))
        assert candidate.getpixel((1300, y)) == (255, 255, 255)
    assert max(diffs) <= 25
    return max(diffs)
print('Maximum sampled channel difference:', check(preview))
for x in (60, 660, 1393):
    altered = preview.copy(); altered.paste((255, 255, 255), (x, 280, x + 300, 780))
    try:
        check(altered)
    except AssertionError:
        continue
    raise AssertionError(f'blanked image at x={x} passed')
for name in ('decoded-sequential.ppm', 'decoded12.ppm', 'decoded-sequential.png', 'decoded-progressive.png', 'resolved-preview.svg', 'resolved-preview.png', 'resolved-preview-report.json'):
    print(digest((work / name).read_bytes()), name)
