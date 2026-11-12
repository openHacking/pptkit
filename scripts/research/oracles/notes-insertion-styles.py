"""Independent source/empty-mark/actual-typing native admission checks."""
import importlib.util
import json
import math
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as ET

import fitz
from PIL import Image, ImageDraw

spec = importlib.util.spec_from_file_location('notes_pages', Path(__file__).with_name('native-notes-pages.py'))
base = importlib.util.module_from_spec(spec)
spec.loader.exec_module(base)
NS, sha, payload, root, edge, raster, fonts = base.NS, base.sha, base.payload, base.root, base.edge, base.raster, base.font_programs
font_spec = importlib.util.spec_from_file_location('slide_fonts', Path(__file__).with_name('native-notes-theme-slides.py'))
font_bounds = importlib.util.module_from_spec(font_spec)
font_spec.loader.exec_module(font_bounds)


def signature(n, character=False):
    attrs = dict(n.attrib)
    if character:
        attrs.pop('dirty', None)
    return (('character-properties' if character else n.tag), tuple(sorted(attrs.items())), n.text or '',
            tuple(signature(c) for c in n))


def owners(file):
    parts = payload(file)
    presentation = root(parts, 'ppt/presentation.xml')
    rows = []
    for identity in presentation.findall('p:sldIdLst/p:sldId', NS):
        slide = edge(parts, 'ppt/presentation.xml', 'slide', identity.get('{' + NS['r'] + '}id'))
        note = edge(parts, slide, 'notesSlide')
        assert edge(parts, note, 'slide') == slide
        doc = root(parts, note)
        bodies = [s for s in doc.findall('p:cSld/p:spTree/p:sp', NS)
                  if s.find('p:nvSpPr/p:nvPr/p:ph', NS).get('type') == 'body']
        assert len(bodies) == 1
        paragraphs = bodies[0].findall('p:txBody/a:p', NS)
        assert len(paragraphs) == 1
        label = ''.join(root(parts, slide).itertext())
        rows.append({'slide': slide, 'notes': note, 'label': label, 'paragraph': paragraphs[0]})
    assert len(rows) == 4 and len({r['notes'] for r in rows}) == 4
    return rows


def package(file, cases, expected_properties, stage):
    rows = owners(file)
    records = []
    for i, row in enumerate(rows):
        p = row['paragraph']
        assert row['label'] == 'Insertion ' + cases[i]['id'], 'Wrong notes page identity'
        end = p.find('a:endParaRPr', NS)
        runs = p.findall('a:r', NS)
        if stage == 'empty':
            assert end is not None
            assert signature(end, True) == expected_properties[i], 'Insertion mark style changed'
            assert not runs, 'Unexpected native characters in empty paragraph'
        else:
            # The captured application moves the inherited style onto the new
            # real run and omits the now-redundant end mark in all typed rounds.
            assert end is None, 'Unexpected typed-stage end-mark rewrite'
            assert len(runs) == 1 and runs[0].find('a:t', NS).text == 'Proof', 'Typed text lost/duplicated'
            actual = runs[0].find('a:rPr', NS)
            assert signature(actual, True) == expected_properties[i], 'Typed text did not inherit insertion style'
        records.append({'slide': row['slide'], 'notes': row['notes'], 'label': row['label'],
                        'runCount': len(runs), 'nativeEndPresent': end is not None,
                        'expectedCharacterProperties': expected_properties[i]})
    return records


def actual_ink_bounds(doc, page, traces):
    boxes = []
    # This named PDF uses one font selection per BT/ET block. Bind by actual
    # resource usage, since body and slide image can contain different subsets
    # with the same PostScript name and unrelated glyph numbering.
    blocks = [block for xref in page.get_contents()
              for block in re.findall(rb'\bBT\b(.*?)\bET\b', doc.xref_stream(xref), re.S)]
    all_traces = page.get_texttrace()
    assert len(blocks) == len(all_traces)
    resources = {f[4]: f for f in page.get_fonts()}
    selected = []
    for block, trace in zip(blocks, all_traces):
        names = re.findall(rb'/([A-Za-z0-9]+)\s+[-.0-9]+\s+Tf\b', block)
        assert len(names) == 1
        font = resources[names[0].decode('ascii')]
        assert font[3].split('+')[-1] == trace['font']
        selected.append((trace, font))
    for trace in traces:
        assert trace['dir'] == (1, 0)
        matches = [font for actual, font in selected if actual == trace]
        assert len(matches) == 1
        data = doc.extract_font(matches[0][0])[3]
        for char in trace['chars']:
            units, (x0, y0, x1, y1) = font_bounds.true_type_bounds(data, char[1])
            scale = trace['size'] / units
            x, y = char[2]
            boxes.append((x + x0 * scale, y - y1 * scale, x + x1 * scale, y - y0 * scale))
    return (min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes))


def masked(page, bounds):
    pix = page.get_pixmap(alpha=False)
    image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
    if bounds:
        ImageDraw.Draw(image).rectangle((math.floor(bounds[0]) - 1, math.floor(bounds[1]) - 1,
                                         math.ceil(bounds[2]) + 1, math.ceil(bounds[3]) + 1), fill='white')
    return image.tobytes()


def pdf(file, cases, stage, reference, empty_reference):
    expected_faces = ['Georgia', 'Arial-BoldItalicMT', 'ArialMT', 'Georgia-Italic']
    records = []
    with fitz.open(file) as doc, fitz.open(reference) as before, fitz.open(empty_reference) as empty:
        assert len(doc) == len(before) == len(empty) == 4
        for i, page in enumerate(doc):
            assert tuple(page.rect) == (0, 0, 595, 842)
            traces = page.get_texttrace()
            image = [t for t in traces if t['bbox'][1] < 400]
            body = [t for t in traces if t['bbox'][1] >= 400]
            assert ''.join(''.join(chr(c[0]) for c in t['chars']) for t in image) == 'Insertion ' + cases[i]['id']
            if stage == 'empty':
                assert not body and raster(page) == raster(before[i]), 'Empty page appearance changed'
            else:
                assert len(body) == 1 and ''.join(chr(c[0]) for c in body[0]['chars']) == 'Proof'
                assert body[0]['font'] == expected_faces[i], 'Wrong actual typed font face'
                bounds = actual_ink_bounds(doc, page, body)
                pixels = raster(page, fitz.Rect(bounds) + (-1, -1, 1, 1))
                assert sum(min(pixels[j:j+3]) < 200 for j in range(0, len(pixels), 3)) > 10, 'Invisible typed glyphs'
                assert masked(page, bounds) == masked(empty[i], bounds), 'Unrelated notes artwork changed'
            assert raster(page) == raster(before[i]), 'Later native raster changed'
            assert traces == before[i].get_texttrace(), 'Later native glyphs changed'
            assert fonts(doc, page) == fonts(before, before[i]), 'Embedded font programs changed'
            records.append({'label': cases[i]['id'], 'rasterSha256': sha(raster(page)), 'fontPrograms': fonts(doc, page),
                            'actualInkBounds': actual_ink_bounds(doc, page, body) if body else None,
                            'typedTrace': [{'font': t['font'], 'size': t['size'], 'color': t['color'], 'bbox': t['bbox']} for t in body]})
    return records


def inspect(directory):
    expected = json.loads((directory / 'expected.json').read_text())
    assert sha((directory / 'source.pptx').read_bytes()) == expected['sourceSha256']
    source_rows = owners(directory / 'source.pptx')
    properties = [signature(r['paragraph'].find('a:endParaRPr', NS), True) for r in source_rows]
    groups = []
    for stage in ['empty', 'typed']:
        rounds = []
        for number in range(1, 4):
            name = f'{stage}-round{number}'
            file = directory / (name + '.pptx')
            rows = package(file, expected['cases'], properties, stage)
            pdf_file = directory / (name + '-notes.pdf')
            reference = directory / ('source-notes.pdf' if stage == 'empty' else 'typed-round1-notes.pdf')
            pages = pdf(pdf_file, expected['cases'], stage, reference, directory / 'empty-round3-notes.pdf')
            rounds.append({'name': name, 'sha256': sha(file.read_bytes()), 'pdfSha256': sha(pdf_file.read_bytes()), 'rows': rows, 'pages': pages})
        groups.append({'stage': stage, 'rounds': rounds})
    negatives = []
    negative_directory = directory / 'negative-controls'
    negative_directory.mkdir(exist_ok=True)
    for name, stage, index, target, field, value in [
        ('empty-size', 'empty', 1, 'a:endParaRPr', 'sz', '1500'),
        ('empty-language', 'empty', 1, 'a:endParaRPr', 'lang', 'en-US'),
        ('empty-family', 'empty', 1, 'a:endParaRPr/a:latin', 'typeface', 'Arial'),
        ('empty-bold', 'empty', 2, 'a:endParaRPr', 'b', '0'),
        ('typed-family', 'typed', 1, 'a:r/a:rPr/a:latin', 'typeface', 'Arial'),
        ('typed-color', 'typed', 1, 'a:r/a:rPr/a:solidFill/a:srgbClr', 'val', 'FFFFFF'),
        ('typed-italic', 'typed', 2, 'a:r/a:rPr', 'i', '0'),
        ('typed-text', 'typed', 1, 'a:r/a:t', None, 'Wrong'),
    ]:
        parts = payload(directory / (stage + '-round1.pptx'))
        part = f'ppt/notesSlides/notesSlide{index}.xml'
        node = root(parts, part)
        body = next(s for s in node.findall('p:cSld/p:spTree/p:sp', NS)
                    if s.find('p:nvSpPr/p:nvPr/p:ph', NS).get('type') == 'body')
        selected = body.find('p:txBody/a:p/' + target, NS)
        assert selected is not None
        if field is None:
            selected.text = value
        else:
            selected.set(field, value)
        parts[part] = ET.tostring(node, encoding='utf-8', xml_declaration=True)
        file = negative_directory / (name + '.pptx')
        with zipfile.ZipFile(file, 'w', zipfile.ZIP_DEFLATED) as z:
            for owner, data in parts.items():
                z.writestr(owner, data)
        try:
            package(file, expected['cases'], properties, stage)
        except AssertionError as error:
            negatives.append({'name': name, 'file': str(file), 'sha256': sha(file.read_bytes()), 'reason': str(error)})
        else:
            raise AssertionError('Accepted actual insertion-style corruption: ' + name)
    # Keep the old real failed input as an explicit rejection control.
    old = Path('.tmp/authored-notes-pages/run-rM4cWd/wide-round1.pptx')
    old_parts = payload(old)
    old_note = root(old_parts, 'ppt/notesSlides/notesSlide4.xml')
    old_end = old_note.find('.//p:txBody/a:p/a:endParaRPr', NS)
    assert old_end.get('sz') != '1200' and old_end.get('b') != '1', 'Old failure control no longer fails'
    result = {'schemaVersion': 1, 'recordedAt': datetime.now(timezone.utc).isoformat(), 'directory': str(directory),
              'application': {'name': 'Microsoft PowerPoint for Mac', 'version': '16.104'},
              'sourceSha256': expected['sourceSha256'], 'groups': groups,
              'negatives': negatives,
              'oldFailure': {'file': str(old), 'sha256': sha(old.read_bytes()), 'missingSizeAndBold': True},
              'total': {'controls': 4, 'emptyNativeReopenPages': 12, 'typedNativeReopenPages': 12,
                        'actualNativeInsertions': 4, 'sourceFirstOpenPages': 4,
                        'actualPptxCorruptionsRejected': len(negatives)},
              'scope': 'All-empty notes paragraph insertion style only, including final-style choice for two empty runs. Native zero-character run identities/actions are not retained; this does not authorize source style writes or claim general trailing insertion/style/theme equivalence.'}
    Path('docs/testing/notes-insertion-styles-evidence.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps(result['total']))


if __name__ == '__main__':
    inspect(Path(sys.argv[1]).resolve())
