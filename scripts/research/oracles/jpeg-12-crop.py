"""Check one native crop after the named 12-bit preservation replacement."""
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
            crop = picture.find('p:blipFill/a:srcRect', ns)
            rid = picture.find('.//a:blip', ns).attrib[f"{{{ns['r']}}}embed"]
            result.append((tuple(int(offset.attrib[k]) for k in ('x', 'y')), tuple(int(extent.attrib[k]) for k in ('cx', 'cy')), {} if crop is None else crop.attrib, archive.read(targets[rid])))
        return result

before = inventory(work / 'preserved-native3.pptx')
after = inventory(work / 'crop.pptx')
assert len(before) == len(after) == 3 and before[:2] == after[:2]
assert before[2][:3] == ((8001000, 1710634), (3302000, 3302000), {})
assert after[2][:3] == ((8845826, 1710634), (2457174, 3302000), {'l': '25585'})
assert before[2][3] == after[2][3] == (root / 'tests/fixtures/jpeg-12bit-progressive.jpg').read_bytes()
first = Image.open(work / 'preserved-reopen3.png').convert('RGB')
reopen = Image.open(work / 'crop-reopen.png').convert('RGB')
def check(candidate):
    assert candidate.size == first.size == (1920, 1080)
    for y in (320, 520, 720):
        for x in (110, 310, 510, 710, 910, 1110, 1430, 1550, 1700):
            assert first.getpixel((x, y)) == candidate.getpixel((x, y))
        assert candidate.getpixel((1300, y)) == (255, 255, 255)
check(reopen)
for x in (60, 660, 1393):
    altered = reopen.copy(); altered.paste((255, 255, 255), (x, 280, x + 300, 780))
    try:
        check(altered)
    except AssertionError:
        continue
    raise AssertionError(f'blanked image at x={x} passed')
for name in ('preserved-native3.pptx', 'crop.pptx', 'crop-selected.png', 'crop-reopen.png'):
    path = work / name
    print(hashlib.sha256(path.read_bytes()).hexdigest(), path.relative_to(root))
