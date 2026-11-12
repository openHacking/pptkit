"""Independent ZIP/XML and native PDF oracle for the rect-only notes repair.

Consumes captured PowerPoint files. Does not launch Office, regenerate fixtures,
or infer complete notes-page semantics from this named source-assembly control.
"""
import hashlib
import json
import posixpath
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as ET

import fitz

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
IMAGE = fitz.Rect(25, 35, 570, 300)
REST = fitz.Rect(0, 300, 595, 842)


def sha(data):
    return hashlib.sha256(data).hexdigest()


def payload(file):
    with zipfile.ZipFile(file) as z:
        assert len(z.namelist()) == len(set(z.namelist())), 'Duplicate ZIP names'
        return {name: z.read(name) for name in z.namelist()}


def root(parts, part):
    return ET.fromstring(parts[part])


def edge(parts, owner, kind, identity=None):
    folder, name = posixpath.split(owner)
    rels = root(parts, folder + '/_rels/' + name + '.rels')
    matches = [r for r in rels if r.get('Type', '').endswith('/' + kind)
               and (identity is None or r.get('Id') == identity)]
    assert len(matches) == 1, ('Missing/ambiguous edge', owner, kind)
    assert matches[0].get('TargetMode') != 'External', 'External notes relationship'
    target = posixpath.normpath(posixpath.join(folder, matches[0].get('Target')))
    assert target in parts, ('Missing target', target)
    return target


def image_shape(node):
    matches = [s for s in node.findall('p:cSld/p:spTree/p:sp', NS)
               if s.find('p:nvSpPr/p:nvPr/p:ph', NS) is not None
               and s.find('p:nvSpPr/p:nvPr/p:ph', NS).get('type') == 'sldImg']
    assert len(matches) == 1, 'Missing/ambiguous slide-image placeholder'
    return matches[0]


def package_controls(file, expected_labels=None):
    parts = payload(file)
    presentation = root(parts, 'ppt/presentation.xml')
    rows = []
    for item in presentation.findall('p:sldIdLst/p:sldId', NS):
        slide = edge(parts, 'ppt/presentation.xml', 'slide', item.get('{' + NS['r'] + '}id'))
        note = edge(parts, slide, 'notesSlide')
        assert edge(parts, note, 'slide') == slide, 'Wrong notes-to-slide backlink'
        master = edge(parts, note, 'notesMaster')
        geom = image_shape(root(parts, master)).find('p:spPr/a:prstGeom', NS)
        assert geom is not None and geom.get('prst') == 'rect', 'Missing master rect'
        assert geom.find('a:avLst', NS) is not None, 'Missing rect adjustment list'
        image_shape(root(parts, note))
        shapes = root(parts, slide).findall('p:cSld/p:spTree/p:sp', NS)
        labels = [''.join(s.itertext()) for s in shapes
                  if s.find('p:nvSpPr/p:cNvPr', NS).get('name') == 'SlideLabel']
        assert len(labels) == 1 and labels[0].startswith('notes-A-'), 'Wrong slide identity'
        rows.append({'slide': slide, 'notes': note, 'master': master, 'label': labels[0]})
    assert len({r['slide'] for r in rows}) == len(rows), 'Aliased slide parts'
    assert len({r['notes'] for r in rows}) == len(rows), 'Aliased notes parts'
    if expected_labels is None:
        assert len(rows) == 15 and len({r['label'] for r in rows}) == 15, 'Lost/duplicated pages'
    else:
        assert [r['label'] for r in rows] == expected_labels, 'Wrong copied-page order'
    return rows


def raster(page, clip=None):
    pix = page.get_pixmap(matrix=fitz.Matrix(1, 1), colorspace=fitz.csRGB, alpha=False, clip=clip)
    return pix.samples


def font_programs(doc, page):
    return sorted((font[3], font[2], sha(doc.extract_font(font[0])[3]))
                  for font in page.get_fonts(full=True))


def check_pdf(file, controls, baseline, compare_image=True):
    result = []
    with fitz.open(file) as doc, fitz.open(baseline) as before:
        assert len(doc) == len(before) == len(controls), 'Wrong native page count'
        for i, row in enumerate(controls):
            page, old = doc[i], before[i]
            assert tuple(page.rect) == tuple(old.rect) == (0, 0, 595, 842), 'Wrong notes size'
            image_text = ''.join(''.join(chr(c[0]) for c in t['chars'])
                                 for t in page.get_texttrace() if t['bbox'][1] < 300)
            assert row['label'] in image_text, ('Wrong/blank slide image', i, image_text)
            ink = raster(page, IMAGE)
            assert sum(min(ink[j:j+3]) < 200 for j in range(0, len(ink), 3)) > 100, 'No image ink'
            assert raster(page, REST) == raster(old, REST), ('Changed body/decoration raster', i)
            # Adding a drawing changes PDF paint sequence numbers; compare every
            # glyph/font/geometry field, and independently require exact pixels.
            body = lambda p: [{k: v for k, v in t.items() if k != 'seqno'}
                              for t in p.get_texttrace() if t['bbox'][1] >= 300]
            assert body(page) == body(old), ('Changed body glyphs', i)
            if compare_image:
                assert raster(page) == raster(old), ('Changed native full-page raster', i)
                assert page.get_texttrace() == old.get_texttrace(), ('Changed full-page glyphs', i)
                assert font_programs(doc, page) == font_programs(before, old), ('Changed embedded fonts', i)
            result.append({'label': row['label'], 'rasterSha256': sha(raster(page)),
                           'bodyDecorationSha256': sha(raster(page, REST)),
                           'imageInkPixels': sum(min(ink[j:j+3]) < 200 for j in range(0, len(ink), 3)),
                           'embeddedFontPrograms': font_programs(doc, page)})
    return result


def inspect(directory):
    original = payload(directory / 'patched-original.pptx')
    fixed = payload(directory / 'patched-master-rect.pptx')
    assert original.keys() == fixed.keys(), 'Rect experiment changed package parts'
    master = 'ppt/notesMasters/notesMaster1.xml'
    assert [p for p in original if original[p] != fixed[p]] == [master], 'Non-master payload changed'
    old, new = root(original, master), root(fixed, master)
    shape = image_shape(new)
    props = shape.find('p:spPr', NS)
    rect = props.find('a:prstGeom', NS)
    assert rect is not None and rect.get('prst') == 'rect'
    props.remove(rect)
    assert ET.tostring(old) == ET.tostring(new), 'Rect was not the only semantic change'
    controls = package_controls(directory / 'patched-master-rect.pptx')
    first = check_pdf(directory / 'patched-master-rect-notes.pdf', controls,
                      directory / 'patched-original-notes.pdf', False)
    rounds = []
    for n in range(1, 4):
        file = directory / f'rect-round{n}.pptx'
        rows = package_controls(file)
        assert [r['label'] for r in rows] == [r['label'] for r in controls], 'Reordered notes'
        pdf = directory / f'rect-round{n}-notes.pdf'
        pages = check_pdf(pdf, rows, directory / 'patched-master-rect-notes.pdf')
        rounds.append({'name': file.name, 'pptxSha256': sha(file.read_bytes()),
                       'pdfSha256': sha(pdf.read_bytes()), 'controls': rows, 'pages': pages})
    copied = []
    labels = [controls[0]['label']] + [r['label'] for r in controls]
    for n in range(1, 4):
        file = directory / f'rect-copy-round{n}.pptx'
        rows = package_controls(file, labels)
        pdf = directory / f'rect-copy-round{n}-notes.pdf'
        pages = check_pdf(pdf, rows, directory / 'rect-copy-round1-notes.pdf')
        with fitz.open(pdf) as doc, fitz.open(directory / 'rect-round3-notes.pdf') as before:
            for index, page in enumerate(doc):
                source_index = max(0, index - 1)
                assert raster(page) == raster(before[source_index]), ('Copied page appearance changed', index)
                assert page.get_texttrace() == before[source_index].get_texttrace(), ('Copied page glyphs changed', index)
                assert font_programs(doc, page) == font_programs(before, before[source_index]), ('Copied page font programs changed', index)
        copied.append({'name': file.name, 'pptxSha256': sha(file.read_bytes()),
                       'pdfSha256': sha(pdf.read_bytes()), 'controls': rows, 'pages': pages})
    rejected = []
    def reject(label, action):
        try:
            action()
        except (AssertionError, KeyError):
            rejected.append(label)
        else:
            raise AssertionError('Accepted negative: ' + label)
    for name in ['patched-original', 'patched-no-image-text', 'patched-default-image-index', 'patched-native-image']:
        reject(name, lambda name=name: check_pdf(directory / (name + '-notes.pdf'), controls,
                                                directory / 'patched-master-rect-notes.pdf'))
    swapped = controls.copy()
    swapped[0], swapped[1] = swapped[1], swapped[0]
    reject('wrong-page-correspondence', lambda: check_pdf(directory / 'rect-round3-notes.pdf', swapped,
                                                        directory / 'patched-master-rect-notes.pdf'))
    reject('missing-master-rect', lambda: package_controls(directory / 'patched-original.pptx'))
    files = sorted(p for p in directory.iterdir() if p.suffix in ('.pptx', '.pdf') and not p.name.startswith('~$'))
    return {'schemaVersion': 1, 'recordedAt': datetime.now(timezone.utc).isoformat(),
            'application': {'name': 'Microsoft PowerPoint for Mac', 'version': '16.104',
                            'operation': 'GUI Save As, close, exact-path reopen, portrait Notes print'},
            'directory': str(directory.resolve()),
            'files': [{'name': p.name, 'sha256': sha(p.read_bytes())} for p in files],
            'rectOnlyChangeVerified': True, 'firstOpenPages': first, 'rounds': rounds, 'copiedSlideRounds': copied,
            'rejectedControls': rejected, 'total': {'sourcePages': 15, 'nativeReopenedPages': 45,
                                                 'copiedSlideNativeReopenedPages': 48,
                                                 'negativeControlsRejected': len(rejected)},
            'scope': 'Named A source assembly and native first-slide duplication only. Complete fresh-author notes pages, fields, B/override image admission, Windows and source style writes remain open.'}


if __name__ == '__main__':
    directory = Path(sys.argv[1]).resolve()
    result = inspect(directory)
    output = Path('docs/testing/native-notes-pages-evidence.json')
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps(result['total']))
