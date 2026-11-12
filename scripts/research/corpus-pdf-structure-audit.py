#!/usr/bin/env python3
"""Compare PDF structures from the retained headless render experiment."""
import hashlib
import json
import sys
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[2]


def sha(data):
    return hashlib.sha256(data).hexdigest()


def spans(page):
    blocks = page.get_text('dict', flags=fitz.TEXTFLAGS_TEXT)['blocks']
    return [span for block in blocks if block['type'] == 0 for line in block['lines'] for span in line['spans']]


def rect_delta(a, b):
    return max(abs(x - y) for x, y in zip(a, b))


def image_inventory(page):
    images = []
    for image in page.get_images(full=True):
        xref = image[0]
        bitmap = fitz.Pixmap(page.parent, xref)
        images.append({'width': bitmap.width, 'height': bitmap.height, 'components': bitmap.n, 'pixelsSha256': sha(bitmap.samples), 'placements': [list(rect) for rect in page.get_image_rects(xref)]})
    return images


def compare_pages(a, b):
    left, right = spans(a), spans(b)
    drawing_a, drawing_b = a.get_drawings(), b.get_drawings()
    text = {'spanCountSource': len(left), 'spanCountSaved': len(right), 'extractedTextEqual': a.get_text('text') == b.get_text('text')}
    if len(left) == len(right):
        text['ordinalSpans'] = [{'index': i, 'textEqual': x['text'] == y['text'], 'fontEqual': x['font'] == y['font'], 'sizeDelta': abs(x['size'] - y['size']), 'bboxMaxDelta': rect_delta(x['bbox'], y['bbox']), 'originMaxDelta': rect_delta(x['origin'], y['origin'])} for i, (x, y) in enumerate(zip(left, right))]
    drawing = {'countSource': len(drawing_a), 'countSaved': len(drawing_b)}
    if len(drawing_a) == len(drawing_b):
        drawing['ordinalPaths'] = [{'index': i, 'equal': x == y, 'rectMaxDelta': rect_delta(x['rect'], y['rect']), 'differentFields': sorted(k for k in set(x) | set(y) if x.get(k) != y.get(k))} for i, (x, y) in enumerate(zip(drawing_a, drawing_b))]
    return {'text': text, 'drawings': drawing, 'sourceImages': image_inventory(a), 'savedImages': image_inventory(b)}


def main():
    source = ROOT / 'docs/testing/corpus-headless-render-audit.json'
    manifest = json.loads(source.read_text())
    directory = ROOT / manifest['directory']
    fixtures = []
    for item in manifest['fixtures']:
        name = item['fixture']
        left, right = directory / 'source' / f'{name}.pdf', directory / 'saved' / f'{name}.pdf'
        if sha(left.read_bytes()) != item['sourcePdfSha256'] or sha(right.read_bytes()) != item['savedPdfSha256']:
            raise ValueError(f'PDF hash mismatch: {name}')
        with fitz.open(left) as a, fitz.open(right) as b:
            fixtures.append({'fixture': name, 'sourcePdfSha256': item['sourcePdfSha256'], 'savedPdfSha256': item['savedPdfSha256'], 'sourcePageCount': len(a), 'savedPageCount': len(b), 'pages': [compare_pages(a[i], b[i]) for i in range(min(len(a), len(b)))]})
    result = {'schemaVersion': 1, 'sourceEvidence': str(source.relative_to(ROOT)), 'sourceEvidenceSha256': sha(source.read_bytes()), 'scriptSha256': sha(Path(__file__).read_bytes()), 'pymupdfVersion': fitz.VersionBind, 'scope': 'Ordinal PDF structure comparison only; equal ordinals are not independently established object identities and extracted text is not proof of visual fidelity.', 'fixtures': fixtures}
    output = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / '.tmp/corpus-pdf-structure-audit.json'
    output.write_text(json.dumps(result, indent=2) + '\n')
    for item in fixtures:
        if not item['pages']:
            print(item['fixture'], 'no paired pages')
            continue
        page = item['pages'][0]
        print(item['fixture'], 'spans', page['text']['spanCountSource'], page['text']['spanCountSaved'], 'drawing changes', sum(not x['equal'] for x in page['drawings'].get('ordinalPaths', [])), 'images equal', page['sourceImages'] == page['savedImages'])


if __name__ == '__main__':
    main()
