"""Verify one PowerPoint shape-text edit, saved package, reopening and print."""
import hashlib
import json
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
import fitz

root = Path(__file__).resolve().parents[2]
base = root/'.tmp/native-tab-stops/powerpoint-shape-j8Dn0A'
evidence = Path(next((arg for arg in sys.argv[1:] if arg != '--record'), root/'docs/testing/tab-shape-native-edit-powerpoint-evidence.json'))
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
prior = root/'docs/testing/tab-shape-reopened-powerpoint-evidence.json'
reference = json.loads(prior.read_text())
source, edited = base/'shape-round-3.pptx', base/'shape-native-edit.pptx'
pdf_before, pdf_after = base/'shape-round-3.pdf', base/'shape-native-edit.pdf'
assert sha(source) == reference['thirdSavedPptxSha256'] and sha(pdf_before) == reference['reopenedPdfSha256']
a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
p = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
cases = [('left', 'l', ['I', 'WIDE', 'MIXED']), ('center', 'ctr', ['I', 'WIDE', 'MIXED']), ('right', 'r', ['I', 'WIDE', 'MIXED']), ('decimal', 'dec', ['1.2', '123.45', '0.678'])]
with zipfile.ZipFile(edited) as package:
    assert package.testzip() is None
    for page_index, (name, alignment, samples) in enumerate(cases, 1):
        slide = ET.fromstring(package.read(f'ppt/slides/slide{page_index}.xml'))
        shapes = {shape.find(p+'nvSpPr/'+p+'cNvPr').get('name'): shape for shape in slide.iter(p+'sp')}
        for i, sample in enumerate(samples):
            shape = shapes[f'{name}-{i}']; para = shape.find(p+'txBody/'+a+'p')
            assert shape.find(p+'nvSpPr/'+p+'cNvSpPr').get('txBox') is None
            assert shape.find(p+'spPr/'+a+'solidFill/'+a+'srgbClr').get('val') == 'F2F5FA'
            assert para.find(a+'pPr/'+a+'tabLst/'+a+'tab').attrib == {'pos':'3556000','algn':alignment}
            assert ''.join(node.text or '' for node in para.iter(a+'t')) == 'P\t'+sample+('X' if name == 'left' and i == 1 else '')
ax = base/'shape-native-edit-reopen.ax.txt'; state = ax.read_text()
assert 'shape-native-edit.pptx' in state and 'Help: left-1, Value: P\tWIDEX' in state
with fitz.open(pdf_before) as before, fitz.open(pdf_after) as after:
    assert len(before) == len(after) == 4
    for index, (first, second) in enumerate(zip(before, after)):
        first_words, second_words = first.get_text('words'), second.get_text('words')
        assert len(first_words) == len(second_words)
        for old, new in zip(first_words, second_words):
            if index == 0 and old[4] == 'WIDE':
                assert new[4] == 'WIDEX' and old[:2] == new[:2] and old[3] == new[3] and new[2] > old[2]
                changed = {'before': list(old[:5]), 'after': list(new[:5])}
            else:
                assert old == new
        def fills(page):
            return [(tuple(d['rect']), d['fill']) for d in page.get_drawings() if d['fill'] and all(abs(d['fill'][i] - v/255) < .01 for i, v in enumerate((242,245,250)))]
        assert len(fills(first)) == 3 and fills(first) == fills(second)
assert changed['before'][4] == 'WIDE'
result = {'scope':'One appended X in one left-aligned shape-text sample, saved and reopened in PowerPoint. All 12 saved sample texts, fills and tab declarations checked; four printed pages compared, including unchanged text bounds and edited word start. Other edit operations, shape transforms and broader native parity remain unverified. Capture hashes do not independently establish provenance.', 'priorEvidenceSha256':sha(prior), 'editedPptxSha256':sha(edited), 'editedPdfSha256':sha(pdf_after), 'reopenedAxSha256':sha(ax), 'reopenedScreenshotSha256':sha(base/'shape-native-edit-reopen.png'), 'verifierSha256':sha(Path(__file__)), 'changedWord':changed}
if '--record' in sys.argv: evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else: assert json.loads(evidence.read_text()) == result
print(json.dumps({'verified':True,'slides':4,'samples':12,'edited':1}))
