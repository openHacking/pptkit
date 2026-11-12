"""Verify retained PowerPoint first-open, print, and save/reopen shape-text observations."""
import hashlib
import json
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
import fitz

root = Path(__file__).resolve().parents[2]
base = root / '.tmp/native-tab-stops/powerpoint-shape-j8Dn0A'
evidence = Path(next((arg for arg in sys.argv[1:] if arg != '--record'), root / 'docs/testing/tab-shape-powerpoint-evidence.json'))
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
p = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
cases = [('left', 'l', ['I', 'WIDE', 'MIXED']), ('center', 'ctr', ['I', 'WIDE', 'MIXED']), ('right', 'r', ['I', 'WIDE', 'MIXED']), ('decimal', 'dec', ['1.2', '123.45', '0.678'])]
source, pdf = base / 'shapes.pptx', base / 'shapes.pdf'
assert sha(source) == json.loads((base / 'manifest.json').read_text())['sha256']
with zipfile.ZipFile(source) as package:
    assert package.testzip() is None
    for index, (name, alignment, samples) in enumerate(cases, 1):
        slide = ET.fromstring(package.read(f'ppt/slides/slide{index}.xml'))
        shapes = {shape.find(p+'nvSpPr/'+p+'cNvPr').get('name'): shape for shape in slide.iter(p+'sp')}
        for number, sample in enumerate(samples):
            shape = shapes[f'{name}-{number}']; para = shape.find(p+'txBody/'+a+'p')
            assert shape.find(p+'nvSpPr/'+p+'cNvSpPr').get('txBox') is None
            assert shape.find(p+'spPr/'+a+'prstGeom').get('prst') == 'rect'
            assert shape.find(p+'spPr/'+a+'solidFill/'+a+'srgbClr').get('val') == 'F2F5FA'
            assert para.find(a+'pPr/'+a+'tabLst/'+a+'tab').attrib == {'pos': '3556000', 'algn': alignment}
            assert [run.find(a+'t').text for run in para.findall(a+'r')] == ['P\t', sample]
records = []
with fitz.open(pdf) as document:
    assert len(document) == 4
    for index, (page, (name, _, samples)) in enumerate(zip(document, cases), 1):
        drawings = page.get_drawings()
        guides = [item for item in drawings if item['fill'] and item['fill'][0] > .7 and item['fill'][1] < .3]
        assert len(guides) == 1
        guide = (guides[0]['rect'].x0 + guides[0]['rect'].x1) / 2
        fills = [item for item in drawings if item['fill'] and all(abs(item['fill'][i] - v/255) < .01 for i, v in enumerate((242,245,250)))]
        assert len(fills) == 3
        words = [word for word in page.get_text('words') if word[4] in {'P', *samples}]
        assert [word[4] for word in words] == [v for sample in samples for v in ('P', sample)]
        dots = [char['bbox'][0] for block in page.get_text('rawdict')['blocks'] if 'lines' in block for line in block['lines'] for span in line['spans'] for char in span['chars'] if char['c'] == '.']
        rows = []
        for number, sample in enumerate(samples):
            prefix, word = words[number*2:number*2+2]
            anchor = word[0] if name == 'left' else (word[0]+word[2])/2 if name == 'center' else word[2] if name == 'right' else dots[number]
            assert abs(anchor-guide) < .25 and abs(prefix[1]-word[1]) < .25
            rows.append({'text': sample, 'prefix': [round(x,4) for x in prefix[:4]], 'word': [round(x,4) for x in word[:4]], 'anchor': round(anchor,4)})
        label = 'first-open' if index == 1 else f'slide-{index}'
        ax = base / f'{label}.ax.txt'; state = ax.read_text()
        assert f'title-{name}' in state and all(f'{name}-{number}' in state for number in range(3))
        records.append({'slide': name, 'guide': round(guide,4), 'rows': rows, 'axSha256': sha(ax), 'screenshotSha256': sha(base / f'{label}.png')})
rounds = []
for number in range(1, 4):
    saved = base / f'shape-round-{number}.pptx'
    with zipfile.ZipFile(saved) as package:
        assert package.testzip() is None
        for index, (name, alignment, samples) in enumerate(cases, 1):
            slide = ET.fromstring(package.read(f'ppt/slides/slide{index}.xml'))
            shapes = {shape.find(p+'nvSpPr/'+p+'cNvPr').get('name'): shape for shape in slide.iter(p+'sp')}
            for item, sample in enumerate(samples):
                shape = shapes[f'{name}-{item}']; para = shape.find(p+'txBody/'+a+'p')
                assert shape.find(p+'nvSpPr/'+p+'cNvSpPr').get('txBox') is None
                assert shape.find(p+'spPr/'+a+'solidFill/'+a+'srgbClr').get('val') == 'F2F5FA'
                assert para.find(a+'pPr/'+a+'tabLst/'+a+'tab').attrib == {'pos': '3556000', 'algn': alignment}
                assert ''.join(node.text or '' for node in para.iter(a+'t')) == 'P\t'+sample
    ax = base / f'round-{number}-reopen.ax.txt'; state = ax.read_text()
    assert f'shape-round-{number}.pptx' in state and 'left-1' in state and 'P\tWIDE' in state
    rounds.append({'pptxSha256': sha(saved), 'axSha256': sha(ax), 'screenshotSha256': sha(base / f'round-{number}-reopen.png')})
version = plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString']
result = {'scope': 'PowerPoint first open and A4 scaled print of 12 shape-text samples; four page AX/screenshot captures and three Save As/close/reopen rounds with XML checks of every sample and first-page window checks. No reopened PDF, native edit or shape transform observation. Screenshots and AX hashes are not independent proof of provenance.', 'applicationVersion': version, 'directory': str(base.relative_to(root)), 'sourceSha256': sha(source), 'pdfSha256': sha(pdf), 'generatorSha256': sha(root/'scripts/prepare-native-shape-tab-stops.mjs'), 'verifierSha256': sha(Path(__file__)), 'slides': records, 'rounds': rounds}
if '--record' in sys.argv: evidence.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n')
else: assert json.loads(evidence.read_text()) == result
print(json.dumps({'verified': True, 'slides': len(records), 'samples': sum(len(item['rows']) for item in records)}))
