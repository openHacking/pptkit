import hashlib
import json
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from PIL import Image

root = Path(__file__).resolve().parents[2]
base = root / '.tmp/native-svg-opacity'
record = root / 'docs/testing/svg-opacity-powerpoint-evidence.json'
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def inspect(path):
    with zipfile.ZipFile(path) as archive: slide = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
    result = []
    for shape in slide.findall('.//p:sp', ns):
        prop = shape.find('p:spPr', ns); off = prop.find('a:xfrm/a:off', ns)
        colors = prop.findall('.//a:srgbClr', ns)
        result.append({'offset': [int(off.get('x')), int(off.get('y'))], 'colors': [{'rgb': color.get('val'), 'alpha': int(color.find('a:alpha', ns).get('val'))} for color in colors]})
    return result
before, after = inspect(base / 'original.pptx'), inspect(base / 'source.pptx')
expected = [[{'rgb': 'FF0000', 'alpha': 26667}], [{'rgb': '00FF00', 'alpha': 40157}], [{'rgb': '0000FF', 'alpha': 25000}]]
assert len(before) == len(after) == 3 and [item['colors'] for item in before] == expected
assert [item['colors'] for item in after] == expected
assert after[0]['offset'][0] > before[0]['offset'][0] and after[0]['offset'][1] == before[0]['offset'][1] and after[1:] == before[1:]
samples = {}
for name in ('first', 'reopen'):
    image = Image.open(base / (name + '.png')).convert('RGB'); assert image.size == (1600, 900)
    red, green, blue, background = [list(image.getpixel(point)) for point in ((580, 400), (860, 275), (1240, 400), (700, 650))]
    assert red[0] > 240 and 150 < red[1] < 220 and 150 < red[2] < 220
    assert green[1] > 240 and 100 < green[0] < 210 and 100 < green[2] < 210
    assert blue[2] > 240 and 150 < blue[0] < 220 and 150 < blue[1] < 220 and background == [255, 255, 255]
    state = (base / (name + '.ax.txt')).read_text(); assert all(f'element-{i}' in state for i in (1, 2, 3))
    samples[name] = {'redFill': red, 'greenStroke': green, 'blueFill': blue, 'background': background}
assert samples['first'] == samples['reopen']
files = ['source.svg', 'original.pptx', 'source.pptx', 'first.png', 'reopen.png', 'first.ax.txt', 'reopen.ax.txt']
data = {'scope': 'PowerPoint first open and one interactive move/save/close/reopen of three transparent fill/stroke samples. Package XML checks colors and alpha; fixed screen samples check color trends and first/reopen equality. Display color management affects RGB; these samples are not an exact sRGB or full-pixel comparison. Captures do not authenticate provenance. Other backgrounds, overlapping objects, transparency editing and further rounds remain unverified.', 'applicationVersion': plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'], 'directory': '.tmp/native-svg-opacity', 'generatorSha256': sha(root / 'scripts/prepare-svg-opacity-native.mjs'), 'verifierSha256': sha(Path(__file__)), 'sha256': {name: sha(base / name) for name in files}, 'before': before, 'after': after, 'samples': samples}
if '--record' in sys.argv: record.write_text(json.dumps(data, indent=2) + '\n')
else: assert json.loads(record.read_text()) == data
print('Verified SVG opacity PowerPoint evidence.')
