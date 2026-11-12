"""Independently inspect a named Chromium PDF font-selection experiment."""
import hashlib
import json
import sys
from pathlib import Path
import fitz

if len(sys.argv) != 2:
    raise SystemExit('Usage: python3 pdf-font.py CAPTURE_DIRECTORY')
root = Path(sys.argv[1])
manifest = json.loads((root / 'evidence.json').read_text())
digest = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
assert digest(Path('scripts/research/pdf-font-observation.mjs')) == manifest['producerSha256']
assert digest(Path('packages/pdf-renderer/src/rendering/print.ts')) == manifest['adapterSourceSha256']
assert digest(Path('packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2')) == manifest['fontSha256']
assert manifest['foreignObject'] == [False, True]
assert manifest['corrupt'] == 'pdf-font-unavailable'
assert digest(root / 'coverage.pdf') == manifest['coverage']['sha256']
def inspect(doc):
    assert len(doc) == 2 and [page.get_text().strip() for page in doc] == ['iiii WWWW 12345', 'Linked 12345']
    pages = []
    for page in doc:
        spans = [span for block in page.get_text('dict')['blocks'] if 'lines' in block for line in block['lines'] for span in line['spans']]
        assert len(spans) == 1
        fonts = page.get_fonts(full=True)
        assert len(fonts) == 1 and len(doc.extract_font(fonts[0][0])[3]) > 0
        pages.append({'font': spans[0]['font'], 'width': spans[0]['bbox'][2] - spans[0]['bbox'][0]})
    assert [link.get('uri') for link in doc[1].get_links()] == ['https://example.com/font']
    return pages

observations = {}
for item in manifest['results']:
    path = root / (item['name'] + '.pdf')
    assert digest(path) == item['sha256']
    with fitz.open(path) as doc:
        observations[item['name']] = inspect(doc)
assert all(page['font'] == 'Times-Roman' for page in observations['plain'])
assert all(page['font'] == 'RobotoMono-Regular' for page in observations['embedded'])
assert all(abs(a['width'] - b['width']) > 5 for a, b in zip(observations['plain'], observations['embedded']))
with fitz.open(root / 'coverage.pdf') as coverage:
    assert len(coverage) == 1 and coverage[0].get_text().strip() == 'ABC 汉字'
    spans = [span for block in coverage[0].get_text('dict')['blocks'] if 'lines' in block for line in block['lines'] for span in line['spans']]
    assert ''.join(span['text'] for span in spans) == 'ABC 汉字'
    assert spans[0]['font'] == 'RobotoMono-Regular' and all(span['font'].startswith('Type3') for span in spans[1:])
    assert any(warning['code'] == 'pdf-font-selection-unverified' for warning in manifest['coverage']['warnings'])
with fitz.open(root / 'embedded.pdf') as modified:
    modified.select([1, 0])
    try:
        inspect(modified)
    except AssertionError:
        pass
    else:
        raise AssertionError('Swapped font sample unexpectedly passed')
with fitz.open(root / 'embedded.pdf') as modified:
    page = modified[1]
    page.delete_link(page.get_links()[0])
    altered = modified.tobytes()
with fitz.open(stream=altered, filetype='pdf') as modified:
    try:
        inspect(modified)
    except AssertionError:
        pass
    else:
        raise AssertionError('Font sample with deleted link unexpectedly passed')
print(json.dumps({'mupdf': fitz.VersionBind, 'pages': observations, 'corrupt': manifest['corrupt'], 'coverageFallback': True, 'negativeControls': 2}))
