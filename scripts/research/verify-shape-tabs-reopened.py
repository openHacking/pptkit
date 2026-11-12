"""Compare PowerPoint's third-reopen shape-tab print with its first-open print."""
import hashlib
import json
import sys
from pathlib import Path
import fitz

root = Path(__file__).resolve().parents[2]
base = root / '.tmp/native-tab-stops/powerpoint-shape-j8Dn0A'
evidence = Path(next((arg for arg in sys.argv[1:] if arg != '--record'), root / 'docs/testing/tab-shape-reopened-powerpoint-evidence.json'))
sha = lambda file: hashlib.sha256(file.read_bytes()).hexdigest()
first, reopened = base/'shapes.pdf', base/'shape-round-3.pdf'
source = root/'docs/testing/tab-shape-powerpoint-evidence.json'
original = json.loads(source.read_text())
assert sha(first) == original['pdfSha256']
assert sha(base/'shape-round-3.pptx') == original['rounds'][2]['pptxSha256']
cases = [('left', ['I', 'WIDE', 'MIXED']), ('center', ['I', 'WIDE', 'MIXED']), ('right', ['I', 'WIDE', 'MIXED']), ('decimal', ['1.2', '123.45', '0.678'])]
observations = []
with fitz.open(first) as initial, fitz.open(reopened) as later:
    assert len(initial) == len(later) == len(cases)
    for page_index, (a, b, (kind, samples)) in enumerate(zip(initial, later, cases), 1):
        assert tuple(a.rect) == tuple(b.rect)
        wa, wb = a.get_text('words'), b.get_text('words')
        assert wa == wb
        selected = [word for word in wb if word[4] in {'P', *samples}]
        assert [word[4] for word in selected] == [value for sample in samples for value in ('P', sample)]
        def shapes(page):
            return [[list(round(v, 5) for v in d['rect']), list(round(v, 5) for v in d['fill'])] for d in page.get_drawings() if d['fill'] and all(abs(d['fill'][i] - v/255) < .01 for i, v in enumerate((242, 245, 250)))]
        fills = shapes(b)
        assert len(fills) == 3 and shapes(a) == fills
        guides = [d for d in b.get_drawings() if d['fill'] and d['fill'][0] > .7 and d['fill'][1] < .3]
        assert len(guides) == 1
        guide = (guides[0]['rect'].x0 + guides[0]['rect'].x1) / 2
        dots = [char['bbox'][0] for block in b.get_text('rawdict')['blocks'] if 'lines' in block for line in block['lines'] for span in line['spans'] for char in span['chars'] if char['c'] == '.']
        rows = []
        for i, sample in enumerate(samples):
            prefix, word = selected[2*i:2*i+2]
            anchor = word[0] if kind == 'left' else (word[0]+word[2])/2 if kind == 'center' else word[2] if kind == 'right' else dots[i]
            assert abs(anchor-guide) < .25 and abs(prefix[1]-word[1]) < .25
            rows.append({'text': sample, 'prefix': [round(v,4) for v in prefix[:4]], 'word': [round(v,4) for v in word[:4]], 'anchor': round(anchor,4)})
        observations.append({'page': page_index, 'kind': kind, 'guide': round(guide,4), 'fills': fills, 'rows': rows})
result = {'scope': 'Four-page A4 print after third save/close/reopen compared with first-open PowerPoint print: text word geometry, three shape fills and tab anchors. No interactive shape edit or transformed-shape claim. PDF hashes do not independently prove application provenance.', 'firstEvidenceSha256': sha(source), 'firstPdfSha256': sha(first), 'thirdSavedPptxSha256': sha(base/'shape-round-3.pptx'), 'reopenedPdfSha256': sha(reopened), 'verifierSha256': sha(Path(__file__)), 'pages': observations}
if '--record' in sys.argv:
    evidence.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n')
else:
    assert json.loads(evidence.read_text()) == result
print(json.dumps({'verified': True, 'pages': len(observations), 'samples': sum(len(page['rows']) for page in observations)}))
