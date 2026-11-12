"""Verify the named preservation candidate and one native move/reopen."""
from pathlib import Path
import hashlib
import posixpath
import xml.etree.ElementTree as ET
import zipfile
from PIL import Image

root = Path(__file__).resolve().parents[3]
work = root / '.tmp/native-jpeg-12'
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}

def pictures(package):
    slide = ET.fromstring(package['ppt/slides/slide1.xml'])
    rels = ET.fromstring(package['ppt/slides/_rels/slide1.xml.rels'])
    targets = {item.attrib['Id']: posixpath.normpath(posixpath.join('ppt/slides', item.attrib['Target'])) for item in rels}
    result = []
    for picture in slide.findall('.//p:pic', ns):
        transform = picture.find('p:spPr/a:xfrm', ns)
        off, ext = transform.find('a:off', ns), transform.find('a:ext', ns)
        rid = picture.find('.//a:blip', ns).attrib[f"{{{ns['r']}}}embed"]
        result.append((tuple(int(off.attrib[key]) for key in ('x', 'y')), tuple(int(ext.attrib[key]) for key in ('cx', 'cy')), package[targets[rid]]))
    return result

source, candidate, native = [parts(work / name) for name in ('probe.pptx', 'preserved-before.pptx', 'preserved-native1.pptx')]
assert set(candidate) - set(source) == {'ppt/media/pptkit-image-1.jpg'}
assert not set(source) - set(candidate)
assert {name for name in source if source[name] != candidate[name]} == {'[Content_Types].xml', 'ppt/slides/_rels/slide1.xml.rels', 'ppt/slides/slide1.xml'}
before, replaced, edited = map(pictures, (source, candidate, native))
assert len(before) == len(replaced) == len(edited) == 3
assert before[:2] == replaced[:2] == edited[:2]
assert before[2][:2] == replaced[2][:2] and before[2][2] == (work / 'control.png').read_bytes()
assert replaced[2][0] == (8001000, 1651000) and edited[2][0] == (7941366, 1651000)
assert replaced[2][1:] == edited[2][1:]
assert replaced[2][2] == (root / 'tests/fixtures/jpeg-12bit-progressive.jpg').read_bytes()

first = Image.open(work / 'preserved-first.png').convert('RGB')
reopen = Image.open(work / 'preserved-reopen.png').convert('RGB')
for y in (320, 520, 720):
    for dx in (50, 250, 450):
        assert max(abs(a - b) for a, b in zip(first.getpixel((660 + dx, y)), first.getpixel((1260 + dx, y)))) <= 2
def check(candidate_image):
    assert first.size == candidate_image.size == (1920, 1080)
    for y in (320, 520, 720):
        for dx in (50, 250, 450):
            for x in (60, 660, 1260):
                old = first.getpixel((x + dx, y))
                new = candidate_image.getpixel((x + dx - (9 if x == 1260 else 0), y))
                assert max(abs(a - b) for a, b in zip(old, new)) <= 3
check(reopen)
for x in (60, 660, 1251):
    altered = reopen.copy(); altered.paste((255, 255, 255), (x, 260, x + 520, 780))
    try:
        check(altered)
    except AssertionError:
        continue
    raise AssertionError(f'blanked image at x={x} passed')
for name, expected_offset, screenshot, shift in [
    ('preserved-native2.pptx', (7941366, 1710634), 'preserved-reopen2.png', (-9, 9)),
    ('preserved-native3.pptx', (8001000, 1710634), 'preserved-reopen3.png', (0, 9)),
]:
    snapshot = pictures(parts(work / name))
    assert snapshot[:2] == before[:2]
    assert snapshot[2][0] == expected_offset and snapshot[2][1:] == replaced[2][1:]
    screen = Image.open(work / screenshot).convert('RGB')
    assert screen.size == first.size
    for y in (320, 520, 720):
        for dx in (50, 250, 450):
            for x in (60, 660, 1260):
                old = first.getpixel((x + dx, y))
                new = screen.getpixel((x + dx + (shift[0] if x == 1260 else 0), y + (shift[1] if x == 1260 else 0)))
                assert max(abs(a - b) for a, b in zip(old, new)) <= 3
for name in ('probe.pptx', 'preserved-before.pptx', 'preserved-native1.pptx', 'preserved-native2.pptx', 'preserved-native3.pptx', 'preserved-first.png', 'preserved-moved.png', 'preserved-reopen.png', 'preserved-reopen2.png', 'preserved-reopen3.png'):
    path = work / name
    print(hashlib.sha256(path.read_bytes()).hexdigest(), path.relative_to(root))
