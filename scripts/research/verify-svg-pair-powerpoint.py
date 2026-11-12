#!/usr/bin/env python3
"""Independently inspect a deliberately inconsistent SVG/PNG pair and its native resave."""
import hashlib
import io
import json
import posixpath
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile
from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
DIR = ROOT / '.tmp/native-svg-pair'
RECORD = ROOT / 'docs/testing/svg-pair-powerpoint-evidence.json'
FILES = ('valid-source.pptx', 'native-round-1.pptx', 'first.ax.txt', 'first.jpg', 'reopen.ax.txt', 'reopen.jpg')
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'svg': 'http://schemas.microsoft.com/office/drawing/2016/SVG/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
EXPECTED_SVG = b'<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><rect width="20" height="20" fill="red"/></svg>'

def inspect(filename):
    with zipfile.ZipFile(DIR / filename) as package:
        rels = ET.fromstring(package.read('ppt/slides/_rels/slide1.xml.rels'))
        targets = {r.get('Id'): posixpath.normpath(posixpath.join('ppt/slides', r.get('Target'))) for r in rels}
        root = ET.fromstring(package.read('ppt/slides/slide1.xml'))
        pics = root.findall('.//p:pic', NS)
        assert len(pics) == 2
        found = {}
        for pic in pics:
            name = pic.find('p:nvPicPr/p:cNvPr', NS).get('name')
            blip = pic.find('p:blipFill/a:blip', NS)
            base = package.read(targets[blip.get(f'{{{NS["r"]}}}embed')])
            image = Image.open(io.BytesIO(base)).convert('RGB')
            extrema = image.getextrema()
            assert all(lo == hi for lo, hi in extrema)
            svg_node = blip.find('.//svg:svgBlip', NS)
            svg = None if svg_node is None else package.read(targets[svg_node.get(f'{{{NS["r"]}}}embed')])
            found[name] = {'size': list(image.size), 'color': [lo for lo, _ in extrema], 'baseSha256': hashlib.sha256(base).hexdigest(), 'svgSha256': None if svg is None else hashlib.sha256(svg).hexdigest()}
            if svg is not None:
                assert svg == EXPECTED_SVG
        assert set(found) == {'SVG red / PNG blue', 'PNG blue control'}
        return found

def check():
    sha = {name: hashlib.sha256((DIR / name).read_bytes()).hexdigest() for name in FILES}
    for fixture, source_file in [("source", FILES[0]), ("resave", FILES[1])]:
        data = (ROOT / f"tests/fixtures/svg-pair-powerpoint-{fixture}.pptx").read_bytes()
        assert hashlib.sha256(data).hexdigest() == sha[source_file]
    source, saved = inspect(FILES[0]), inspect(FILES[1])
    pair, control = 'SVG red / PNG blue', 'PNG blue control'
    assert source[pair]['size'] == [20, 20] and source[pair]['color'] == [0, 0, 255]
    assert saved[pair]['size'] == [40, 40] and saved[pair]['color'] == [255, 0, 0]
    assert source[pair]['svgSha256'] == saved[pair]['svgSha256']
    assert source[control] == saved[control] and source[control]['color'] == [0, 0, 255] and source[control]['svgSha256'] is None
    for stem, document in [('first', FILES[0]), ('reopen', FILES[1])]:
        state = (DIR / f'{stem}.ax.txt').read_text()
        assert f'/native-svg-pair/{document}' in state and pair in state and control in state
        image = Image.open(DIR / f'{stem}.jpg').convert('RGB')
        assert image.size == (1600, 900)
        red, blue = image.getpixel((600, 500)), image.getpixel((1200, 500))
        assert red[0] > 230 and red[1] < 50 and red[2] < 50
        assert blue[2] > 220 and blue[0] < 50 and blue[1] < 60
    return {'scope': 'One deliberately mismatched opaque red SVG/blue PNG pair plus blue control, one PowerPoint save/reopen. Sampled screen colors and package media are checked; UI provenance is observed, and arbitrary SVG, alpha, edits and other versions remain unverified.', 'reopenRounds': 1, 'artifactDirectory': '.tmp/native-svg-pair', 'source': source, 'saved': saved, 'sha256': sha}

result = check()
if sys.argv[1:] == ['--record']:
    RECORD.write_text(json.dumps(result, indent=2, ensure_ascii=False) + '\n')
else:
    assert json.loads(RECORD.read_text()) == result
print(json.dumps({'verified': True, 'reopenRounds': 1}))
