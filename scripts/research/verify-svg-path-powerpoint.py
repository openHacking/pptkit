import hashlib
import json
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from PIL import Image

root = Path(__file__).resolve().parents[2]
base = root / '.tmp/native-svg-path-fixed'
record = root / 'docs/testing/svg-path-powerpoint-evidence.json'
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def inspect(path):
    with zipfile.ZipFile(path) as archive: slide = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
    result = []
    for shape in slide.findall('.//p:sp', ns):
        sp = shape.find('p:spPr', ns); line = sp.find('a:ln', ns); offset = sp.find('a:xfrm/a:off', ns)
        path = sp.find('a:custGeom/a:pathLst/a:path', ns)
        result.append({'offset': [int(offset.get('x')), int(offset.get('y'))], 'cap': line.get('cap'), 'width': int(line.get('w')), 'roundJoin': line.find('a:round', ns) is not None, 'commands': [child.tag.split('}')[-1] for child in path]})
    return result
before, after = inspect(base / 'original.pptx'), inspect(base / 'source.pptx')
assert len(before) == len(after) == 2
assert [item['commands'] for item in before] == [['moveTo', 'lnTo', 'lnTo'], ['moveTo', 'lnTo', 'lnTo', 'lnTo', 'close']]
assert all(item['cap'] == 'rnd' and item['roundJoin'] and item['width'] == 228600 for item in before + after)
assert after[0]['offset'][0] > before[0]['offset'][0] and after[0]['offset'][1] == before[0]['offset'][1]
assert {**after[0], 'offset': before[0]['offset']} == before[0] and after[1] == before[1]
samples = {}
for name in ['first', 'reopen']:
    image = Image.open(base / (name + '.png')).convert('RGB'); assert image.size == (1600, 900)
    red = list(image.getpixel((432, 548))); blue = list(image.getpixel((943, 317))); white = list(image.getpixel((1100, 400)))
    assert red[0] > 200 and red[1] < 60 and red[2] < 60
    assert blue[2] > 200 and blue[0] < 60 and blue[1] < 60
    assert white == [255, 255, 255]
    assert all(f'element-{i}' in (base / (name + '.ax.txt')).read_text() for i in (1, 2))
    samples[name] = {'roundEnd': red, 'closedCorner': blue, 'interior': white}
files = ['source.svg', 'original.pptx', 'source.pptx', 'first.png', 'reopen.png', 'first.ax.txt', 'reopen.ax.txt']
data = {'scope': 'Two PowerPoint paths, one open and one closed, first open and one interactive move/save/close/reopen. Fixed screen pixels sample one cap, one closed corner and the empty interior; package checks bind round caps, round joins, closure commands and moved geometry. UI captures do not authenticate provenance or establish full pixel parity, other path styles or native point editing.', 'applicationVersion': plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString'], 'directory': '.tmp/native-svg-path-fixed', 'generatorSha256': sha(root / 'scripts/prepare-svg-path-native.mjs'), 'verifierSha256': sha(Path(__file__)), 'sha256': {name: sha(base / name) for name in files}, 'before': before, 'after': after, 'samples': samples}
if '--record' in sys.argv: record.write_text(json.dumps(data, indent=2) + '\n')
else: assert json.loads(record.read_text()) == data
print('Verified SVG path PowerPoint evidence.')
