"""Check the retained first-open probe; this cannot witness the UI action."""
from pathlib import Path
import hashlib
import posixpath
import xml.etree.ElementTree as ET
import zipfile
from PIL import Image

root = Path(__file__).resolve().parents[3]
work = root / '.tmp/native-jpeg-12'
expected = [root / 'tests/fixtures/jpeg-12bit-sequential.jpg', root / 'tests/fixtures/jpeg-12bit-progressive.jpg', work / 'control.png']
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
with zipfile.ZipFile(work / 'probe.pptx') as archive:
    slide = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
    relationships = ET.fromstring(archive.read('ppt/slides/_rels/slide1.xml.rels'))
    targets = {item.attrib['Id']: posixpath.normpath(posixpath.join('ppt/slides', item.attrib['Target'])) for item in relationships}
    pictures = slide.findall('.//p:pic', ns)
    assert len(pictures) == 3
    for picture, source in zip(pictures, expected):
        relationship = picture.find('.//a:blip', ns).attrib[f"{{{ns['r']}}}embed"]
        assert archive.read(targets[relationship]) == source.read_bytes()

image = Image.open(work / 'first.png').convert('RGB')
def check_pixels(candidate):
    assert candidate.size == (1920, 1080)
    for y in (320, 520, 720):
        for dx in (50, 250, 450):
            pixels = [candidate.getpixel((x + dx, y)) for x in (60, 660, 1260)]
            assert max(abs(a - b) for a, b in zip(pixels[0], pixels[1])) <= 3
            assert max(abs(a - b) for a, b in zip(pixels[0], pixels[2])) <= 35
    assert candidate.getpixel((110, 320))[0] > candidate.getpixel((110, 320))[2] + 50
    assert candidate.getpixel((510, 720))[2] > candidate.getpixel((510, 720))[0] + 100

check_pixels(image)
for x in (60, 660, 1260):
    altered = image.copy()
    altered.paste((255, 255, 255), (x, 260, x + 520, 780))
    try:
        check_pixels(altered)
    except AssertionError:
        continue
    raise AssertionError(f'blanked image at x={x} passed')
for path in [work / 'probe.pptx', work / 'first.png', *expected]:
    print(hashlib.sha256(path.read_bytes()).hexdigest(), path.relative_to(root))
