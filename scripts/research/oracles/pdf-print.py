"""Inspect one named Chromium PDF print with MuPDF, independently of its producer."""
import hashlib
import json
import sys
from pathlib import Path

import fitz

if len(sys.argv) != 2:
    raise SystemExit('Usage: python3 pdf-print.py CAPTURE_DIRECTORY')
root = Path(sys.argv[1])
manifest = json.loads((root / 'evidence.json').read_text())
digest = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
assert digest(Path('scripts/research/pdf-print-observation.mjs')) == manifest['producerSha256']
assert digest(Path('packages/pdf-renderer/src/rendering/print.ts')) == manifest['adapterSourceSha256']
assert digest(root / 'printed.pdf') == manifest['pdfSha256']
assert digest(root / 'linked.pdf') == manifest['linkedPdfSha256']
assert digest(root / 'adapter.pdf') == manifest['adapterPdfSha256']
assert digest(root / 'source.html') == manifest['htmlSha256']
assert digest(root / 'linked-source.html') == manifest['linkedHtmlSha256']
assert manifest['slideIds'] == ['mixed', 'simple']
assert manifest['foreignObject'] == [True, False]
assert manifest['networkRequests'] == []
assert [warning['code'] for warning in manifest['warnings']] == ['slide-action-not-linked', 'slide-action-not-linked']
assert manifest['linkedWarnings'] == []
expected = [('Bold linked next', ['https://example.com/beta']),
            ('Alpha example', [])]

def verify(document):
    assert len(document) == 2
    observed = []
    for page in document:
        observed.append({'width': page.rect.width, 'height': page.rect.height,
                         'text': page.get_text().strip(),
                         'links': [link.get('uri') for link in page.get_links()]})
    for page, (text, links) in zip(observed, expected):
        assert abs(page['width'] - 400) < 0.2 and abs(page['height'] - 200) < 0.2
        assert page['text'] == text and page['links'] == links
    return observed

original = (root / 'printed.pdf').read_bytes()
with fitz.open(stream=original, filetype='pdf') as document:
    observed = verify(document)

def reject(document):
    try:
        verify(document)
    except AssertionError:
        return
    raise AssertionError('Modified PDF unexpectedly passed verification')

with fitz.open(stream=original, filetype='pdf') as swapped:
    swapped.select([1, 0])
    reject(swapped)
with fitz.open(stream=original, filetype='pdf') as unlinked:
    page = unlinked[0]
    page.delete_link(page.get_links()[0])
    changed = unlinked.tobytes()
with fitz.open(stream=changed, filetype='pdf') as reopened:
    reject(reopened)
with fitz.open(root / 'linked.pdf') as linked:
    assert len(linked) == 2
    assert linked[0].get_text().strip() == expected[0][0]
    assert linked[1].get_text().strip() == expected[1][0]
    links = linked[0].get_links()
    assert len(links) == 3 and links[0].get('uri') == 'https://example.com/beta'
    assert all(link.get('page') == 1 and link.get('nameddest') == manifest['linkedTargetId'] for link in links[1:])
assert [page['slideId'] for page in manifest['adapterPages']] == ['mixed', 'simple']
assert [warning['code'] for warning in manifest['adapterWarnings']] == ['svg-foreign-object-dependency', 'svg-system-font-dependency', 'svg-system-font-dependency', 'pdf-font-selection-unverified', 'pdf-page-size-rounding']
def verify_adapter(adapter):
    assert len(adapter) == 2
    assert [page.get_text().strip() for page in adapter] == [item[0] for item in expected]
    assert all(abs(page.rect.width - 400) < 0.2 and abs(page.rect.height - 200) < 0.2 for page in adapter)
    pixels = adapter[0].get_pixmap(alpha=False)
    assert pixels.pixel(40, 100) == (255, 0, 0) and pixels.pixel(200, 100) == (255, 255, 255)
    links = adapter[0].get_links()
    assert len(links) == 3 and links[0].get('uri') == 'https://example.com/beta'
    assert all(link.get('page') == 1 and link.get('nameddest') == manifest['linkedTargetId'] for link in links[1:])

with fitz.open(root / 'adapter.pdf') as adapter:
    verify_adapter(adapter)
with fitz.open(root / 'adapter.pdf') as swapped:
    swapped.select([1, 0])
    try:
        verify_adapter(swapped)
    except AssertionError:
        pass
    else:
        raise AssertionError('Swapped adapter pages unexpectedly passed')
with fitz.open(root / 'adapter.pdf') as unlinked:
    page = unlinked[0]
    page.delete_link(page.get_links()[0])
    changed = unlinked.tobytes()
with fitz.open(stream=changed, filetype='pdf') as reopened:
    try:
        verify_adapter(reopened)
    except AssertionError:
        pass
    else:
        raise AssertionError('Adapter PDF with a deleted link unexpectedly passed')
with fitz.open(root / 'adapter.pdf') as covered:
    covered[0].draw_rect(fitz.Rect(10, 80, 110, 130), fill=(1, 1, 1), color=(1, 1, 1), overlay=True)
    changed = covered.tobytes()
with fitz.open(stream=changed, filetype='pdf') as reopened:
    try:
        verify_adapter(reopened)
    except AssertionError:
        pass
    else:
        raise AssertionError('Adapter PDF with covered shape unexpectedly passed')
print(json.dumps({'mupdf': fitz.VersionBind, 'pages': observed, 'negativeControls': 5, 'linkedTargetPage': 1, 'adapterPages': len(manifest['adapterPages'])}))
