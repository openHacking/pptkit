import hashlib
import json
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from PIL import Image

root = Path(__file__).resolve().parents[2]
base = root / '.tmp/native-svg-stroke'
out = root / 'docs/testing/svg-stroke-powerpoint-evidence.json'
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def shapes(path):
    with zipfile.ZipFile(path) as archive:
        slide = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
    result = []
    for shape in slide.findall('.//p:sp', ns):
        properties = shape.find('p:spPr', ns)
        transform = properties.find('a:xfrm', ns)
        line = properties.find('a:ln', ns)
        result.append({'offset': [int(transform.find('a:off', ns).get(k)) for k in ('x', 'y')],
                       'extent': [int(transform.find('a:ext', ns).get(k)) for k in ('cx', 'cy')],
                       'width': int(line.get('w')), 'color': line.find('.//a:srgbClr', ns).get('val'),
                       'geometry': properties.find('a:prstGeom', ns).get('prst')})
    return result

before, after = shapes(base / 'original.pptx'), shapes(base / 'source.pptx')
assert len(before) == len(after) == 3
assert [s['geometry'] for s in before] == ['ellipse'] * 3
assert [s['width'] for s in before] == [152400, 114300, 76200]
assert [s['color'] for s in before] == ['FF0000', '0000FF', '00AA00']
assert after[0]['offset'][0] > before[0]['offset'][0]
assert after[0]['offset'][1] == before[0]['offset'][1]
assert {**after[0], 'offset': before[0]['offset']} == before[0]
assert after[1:] == before[1:]

pixels = {}
for name in ('first', 'reopen'):
    image = Image.open(base / (name + '.png')).convert('RGB')
    assert image.size == (1600, 900)
    colors = [list(image.getpixel(point)) for point in ((414, 415), (900, 312), (1286, 235))]
    assert colors[0][0] > 200 and colors[0][1] < 60 and colors[0][2] < 60
    assert colors[1][2] > 200 and colors[1][0] < 60 and colors[1][1] < 60
    assert colors[2][1] > 120 and colors[2][0] < 60 and colors[2][2] < 60
    assert all(image.getpixel(point) == (255, 255, 255) for point in ((542, 415), (900, 415), (1286, 415)))
    pixels[name] = colors
    state = (base / ('original-reopened.ax.txt' if name == 'first' else 'reopen.ax.txt')).read_text()
    assert all(f'element-{i}' in state for i in (1, 2, 3))

files = ['source.svg', 'original.pptx', 'source.pptx', 'first.png', 'reopen.png', 'original-reopened.ax.txt', 'reopen.ax.txt']
evidence = {'scope': 'PowerPoint first open, one interactive right-arrow move of the first circle, save, close and exact reopen; original-copy AX was captured on a later reopen. Three shape XML geometries, stroke widths and colors are checked before and after; fixed UI pixels check stroke colors and empty centers. UI captures do not authenticate provenance or prove complete visual parity. Other shapes, fonts, joins, applications and additional reopen rounds remain unverified.',
            'applicationVersion': plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'],
            'directory': '.tmp/native-svg-stroke', 'generatorSha256': digest(root / 'scripts/prepare-svg-stroke-native.mjs'),
            'verifierSha256': digest(Path(__file__)), 'sha256': {name: digest(base / name) for name in files},
            'before': before, 'after': after, 'sampledPixels': pixels}
if '--record' in sys.argv:
    out.write_text(json.dumps(evidence, indent=2) + '\n')
else:
    assert json.loads(out.read_text()) == evidence
print('Verified SVG stroke PowerPoint evidence.')
