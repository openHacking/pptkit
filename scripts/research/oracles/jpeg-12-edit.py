"""Check one retained PowerPoint image move and save/reopen observation."""
from pathlib import Path
import hashlib
import posixpath
import xml.etree.ElementTree as ET
import zipfile
from PIL import Image

root = Path(__file__).resolve().parents[3]
work = root / '.tmp/native-jpeg-12'
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def inventory(path):
    with zipfile.ZipFile(path) as archive:
        slide = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
        rels = ET.fromstring(archive.read('ppt/slides/_rels/slide1.xml.rels'))
        targets = {rel.attrib['Id']: posixpath.normpath(posixpath.join('ppt/slides', rel.attrib['Target'])) for rel in rels}
        result = []
        for picture in slide.findall('.//p:pic', ns):
            transform = picture.find('p:spPr/a:xfrm', ns)
            offset, extent = transform.find('a:off', ns), transform.find('a:ext', ns)
            rid = picture.find('.//a:blip', ns).attrib[f"{{{ns['r']}}}embed"]
            result.append((tuple(int(offset.attrib[key]) for key in ('x', 'y')), tuple(int(extent.attrib[key]) for key in ('cx', 'cy')), archive.read(targets[rid])))
        return result

before, after = inventory(work / 'probe.pptx'), inventory(work / 'edit.pptx')
assert len(before) == len(after) == 3
assert before[0][0] == (381000, 1651000) and after[0][0] == (440634, 1651000)
assert after[0][1:] == before[0][1:]
assert after[1:] == before[1:]
assert after[0][2] == (root / 'tests/fixtures/jpeg-12bit-sequential.jpg').read_bytes()
assert after[1][2] == (root / 'tests/fixtures/jpeg-12bit-progressive.jpg').read_bytes()
assert after[2][2] == (work / 'control.png').read_bytes()
first = Image.open(work / 'first.png').convert('RGB')
reopen = Image.open(work / 'reopen.png').convert('RGB')
def check_screenshot(candidate):
    assert first.size == candidate.size == (1920, 1080)
    for y in (320, 520, 720):
        for dx in (50, 250, 450):
            for x in (60, 660, 1260):
                old = first.getpixel((x + dx, y))
                new = candidate.getpixel((x + dx + (9 if x == 60 else 0), y))
                assert max(abs(a - b) for a, b in zip(old, new)) <= 2

check_screenshot(reopen)
for x in (69, 660, 1260):
    altered = reopen.copy()
    altered.paste((255, 255, 255), (x, 260, x + 520, 780))
    try:
        check_screenshot(altered)
    except AssertionError:
        continue
    raise AssertionError(f'blanked image at x={x} passed')
for path in (work / 'probe.pptx', work / 'edit.pptx', work / 'moved.png', work / 'reopen.png'):
    print(hashlib.sha256(path.read_bytes()).hexdigest(), path.relative_to(root))
