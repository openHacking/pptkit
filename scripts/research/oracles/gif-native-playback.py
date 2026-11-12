#!/usr/bin/env python3
"""Inspect retained package and screenshot samples; does not automate native UI."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import json
import base64
import xml.etree.ElementTree as ET
from PIL import Image

ROOT = Path(__file__).resolve().parents[3]
BASE = ROOT / '.tmp/run-_2s_4erk'
N = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
def digest(data): return sha256(data).hexdigest()
def check(ok, message):
    if not ok: raise ValueError(message)
def inspect(gif, pptx, captures, screenshots):
    g = Image.open(BytesIO(gif)); check(g.n_frames == 2 and g.info.get('loop') == 0, 'GIF frames/loop')
    frames = []
    for index in range(g.n_frames):
        g.seek(index); frames.append({'rgb': list(g.convert('RGB').getpixel((20, 20))), 'durationMs': g.info.get('duration'), 'disposal': g.disposal_method})
    check(frames == [{'rgb': [255, 0, 0], 'durationMs': 1500, 'disposal': 2}, {'rgb': [0, 0, 255], 'durationMs': 1500, 'disposal': 2}], 'GIF pixel/timing')
    with ZipFile(BytesIO(pptx)) as z:
        names = z.namelist(); check(len(names) == len(set(names)), 'duplicate ZIP entries')
        rels = ET.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))
        images = {r.get('Id'): r.get('Target') for r in rels.findall('rel:Relationship', N) if r.get('Type', '').endswith('/image')}
        slide = ET.fromstring(z.read('ppt/slides/slide1.xml'))
        ids = [blip.get('{%s}embed' % N['r']) for blip in slide.findall('.//p:pic/p:blipFill/a:blip', N)]
        check(len(ids) == 2 and all(i in images for i in ids), 'two image references')
        targets = [images[i] for i in ids]; check(targets == ['../media/motion.gif'] * 2, 'shared target')
        check(z.read('ppt/media/motion.gif') == gif, 'GIF bytes changed')
    check(len(captures) == len(screenshots) == 78, 'capture count')
    colors = []
    for i, (row, image_bytes) in enumerate(zip(captures, screenshots)):
        check(row['file'] == f'show-{i}.png' and row['start'] <= row['end'], 'capture metadata')
        if i: check(row['start'] >= captures[i-1]['end'], 'capture order')
        img = Image.open(BytesIO(image_bytes)).convert('RGB'); check(img.size == (1920,1080), 'screenshot geometry')
        pixels = [img.getpixel((400,500)), img.getpixel((1200,500))]
        def classify(pixel):
            r, g, b = pixel
            return 'red' if r > 200 and g < 80 and b < 80 else 'blue' if b > 200 and r < 80 and g < 80 else 'other'
        kinds = [classify(p) for p in pixels]; check(kinds[0] == kinds[1] and kinds[0] != 'other', 'unmatched sample colors')
        colors.append(kinds[0])
    # The second uninterrupted capture batch is the actual loop observation.
    changes = []
    for i in range(18, 78):
        if i == 18 or colors[i] != colors[i-1]: changes.append({'file': captures[i]['file'], 'color': colors[i]})
    check([item['color'] for item in changes[:3]] == ['red','blue','red'], 'missing red-blue-red sequence')
    return {'sourceFrames': frames, 'sharedImageTargets': targets, 'captureCount': len(captures), 'secondBatchTransitions': changes}

def inspect_browser(gif, svg, timing, screenshots):
    root = ET.fromstring(svg)
    images = root.findall('.//{http://www.w3.org/2000/svg}image')
    check(len(images) == 2, 'preview image count')
    for image in images:
        href = image.get('href', '')
        check(href.startswith('data:image/gif;base64,') and base64.b64decode(href.split(',', 1)[1], validate=True) == gif, 'preview GIF bytes')
    captures = timing['samples']; check(len(captures) == len(screenshots) == 150, 'browser capture count')
    changes = []; previous = None
    for i, (row, data) in enumerate(zip(captures, screenshots)):
        check(row['file'] == f'browser-{i}.png' and row['start'] <= row['end'], 'browser timing')
        if i: check(row['start'] >= captures[i-1]['end'], 'browser order')
        image = Image.open(BytesIO(data)).convert('RGB'); check(image.size == (960, 540), 'browser geometry')
        pixels = [image.getpixel((210,270)), image.getpixel((640,270))]
        check(pixels[0] == pixels[1] and pixels[0] in ((255,0,0), (0,0,255)), 'browser image samples')
        color = 'red' if pixels[0][0] else 'blue'
        if color != previous: changes.append({'file': row['file'], 'color': color}); previous = color
    check([c['color'] for c in changes[:3]] == ['red', 'blue', 'red'], 'browser loop observation')
    return {'browserVersion': timing['browserVersion'], 'captureCount': len(captures), 'transitions': changes}

def main():
    gif = (BASE/'motion.gif').read_bytes(); pptx = (BASE/'candidate.pptx').read_bytes(); producer = json.loads((BASE/'producer.json').read_text()); captures = json.loads((BASE/'capture-times.json').read_text()); shots = [(BASE/row['file']).read_bytes() for row in captures]
    check(digest(gif) == producer['gifSha256'] and digest(pptx) == producer['pptxSha256'], 'producer hash')
    result = inspect(gif, pptx, captures, shots)
    svg = (BASE/'preview.svg').read_bytes(); browser_timing = json.loads((BASE/'browser-times.json').read_text()); browser_shots = [(BASE/row['file']).read_bytes() for row in browser_timing['samples']]
    browser_result = inspect_browser(gif, svg, browser_timing, browser_shots)
    # Deliberate bad observations must not satisfy the evidence checks.
    negative = 0
    for badgif, badpptx, badcaptures, badshots in [(gif[:-1], pptx, captures, shots), (gif, pptx, captures[:-1], shots), (gif, pptx, captures, [shots[0]] * 78)]:
        try: inspect(badgif, badpptx, badcaptures, badshots)
        except Exception: negative += 1
    for badsvg, badshots in [(svg.replace(b'data:image/gif;base64,', b'data:image/png;base64,'), browser_shots), (svg, [browser_shots[0]] * 150)]:
        try: inspect_browser(gif, badsvg, browser_timing, badshots)
        except Exception: negative += 1
    check(negative == 5, 'negative checks')
    out = {'scope': 'PowerPoint 16.104 manual slideshow of one named generated deck; screenshot samples and package checks do not automate or prove UI provenance, exact durations, all loop behavior or other applications.', 'directory': str(BASE.relative_to(ROOT)), 'gifSha256': digest(gif), 'pptxSha256': digest(pptx), 'captureTimesSha256': digest((BASE/'capture-times.json').read_bytes()), 'screenshots': [{'file': row['file'], 'sha256': digest(data)} for row,data in zip(captures,shots)], 'observation': result, 'previewSha256': digest(svg), 'browserTimesSha256': digest((BASE/'browser-times.json').read_bytes()), 'browserScreenshots': [{'file': row['file'], 'sha256': digest(data)} for row,data in zip(browser_timing['samples'],browser_shots)], 'browserObservation': browser_result, 'negativeCasesRejected': negative, 'oracleSha256': digest(Path(__file__).read_bytes()), 'scripts': {str(path.relative_to(ROOT)): digest(path.read_bytes()) for path in [ROOT/'scripts/prepare-gif-native-source.py', ROOT/'scripts/prepare-gif-native-playback.mjs', ROOT/'scripts/inspect-gif-browser-playback.mjs']}}
    (ROOT/'docs/testing/gif-powerpoint-playback-evidence.json').write_text(json.dumps(out, indent=2) + '\n')
    print(json.dumps(result))
if __name__ == '__main__': main()
