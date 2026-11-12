"""Read captured fresh-author PowerPoint notes; no native application automation."""
import importlib.util
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

import fitz

spec = importlib.util.spec_from_file_location('notes_pages', Path(__file__).with_name('native-notes-pages.py'))
base = importlib.util.module_from_spec(spec)
spec.loader.exec_module(base)
NS, sha, payload, root, edge, raster, fonts = base.NS, base.sha, base.payload, base.root, base.edge, base.raster, base.font_programs


def shape(doc, kind):
    return [s for s in doc.findall('p:cSld/p:spTree/p:sp', NS)
            if s.find('p:nvSpPr/p:nvPr/p:ph', NS) is not None
            and s.find('p:nvSpPr/p:nvPr/p:ph', NS).get('type') == kind]


def frame(s):
    props = s.find('p:spPr', NS)
    geom = props.find('a:prstGeom', NS)
    assert geom is not None and geom.get('prst') == 'rect', 'Missing master rect'
    off, ext = props.find('a:xfrm/a:off', NS), props.find('a:xfrm/a:ext', NS)
    return {**{k: int(v) for k, v in off.attrib.items()}, **{k: int(v) for k, v in ext.attrib.items()}}


def package(file, expected, geometry=None):
    parts = payload(file)
    p = root(parts, 'ppt/presentation.xml')
    assert p.find('p:notesSz', NS).attrib == {'cx': '6858000', 'cy': '9144000'}, 'Wrong notesSize'
    master = edge(parts, 'ppt/presentation.xml', 'notesMaster')
    md = root(parts, master)
    assert len(shape(md, 'sldImg')) == len(shape(md, 'body')) == 1
    frames = {kind: frame(shape(md, kind)[0]) for kind in ['sldImg', 'body']}
    if geometry is not None:
        assert frames == geometry, 'Native changed notes frames'
    rows = []
    for i, identity in enumerate(p.findall('p:sldIdLst/p:sldId', NS)):
        slide = edge(parts, 'ppt/presentation.xml', 'slide', identity.get('{' + NS['r'] + '}id'))
        note = edge(parts, slide, 'notesSlide')
        assert edge(parts, note, 'slide') == slide and edge(parts, note, 'notesMaster') == master
        nd = root(parts, note)
        assert len(shape(nd, 'sldImg')) == 1
        bodies = shape(nd, 'body')
        assert len(bodies) == (0 if i == 2 else 1), 'Absent/empty body distinction lost'
        text = ''.join(bodies[0].itertext()) if bodies else None
        assert text == expected[i]['text'], 'Wrong notes content/copy owner'
        empty_style = None
        if i == 3:
            rp = bodies[0].find('p:txBody/a:p/a:r/a:rPr', NS)
            empty_style = rp is not None and rp.get('b') == '1' and rp.get('sz') == '1200'
            if geometry is None:
                assert empty_style, 'Fresh writer lost empty styled run'
        assert ''.join(root(parts, slide).itertext()) == expected[i]['label'], 'Wrong slide text'
        ids = [n.get('id') for n in nd.findall('.//p:cNvPr', NS)]
        assert len(ids) == len(set(ids)), 'Duplicate drawing identities'
        rows.append({'slide': slide, 'notes': note, 'text': text, 'label': expected[i]['label'],
                     'emptyStyledRunRetained': empty_style})
    assert len(rows) == 4 and len({r['notes'] for r in rows}) == 4, 'Lost/aliased notes'
    return {'geometry': frames, 'master': master, 'theme': edge(parts, master, 'theme'), 'rows': rows}


def pdf(file, expected, reference=None):
    pages = []
    with fitz.open(file) as doc:
        assert len(doc) == 4
        old = fitz.open(reference) if reference else None
        try:
            for i, page in enumerate(doc):
                assert tuple(page.rect) == (0, 0, 595, 842), 'Wrong native notes paper size'
                traces = page.get_texttrace()
                text = lambda predicate: ''.join(''.join(chr(c[0]) for c in t['chars']) for t in traces if predicate(t))
                assert text(lambda t: t['bbox'][1] < 400) == expected[i]['label'], 'Blank/wrong slide image'
                assert text(lambda t: t['bbox'][1] >= 400) == (expected[i]['text'] or ''), 'Wrong/extra notes body'
                ink = {}
                for role, selected in [('image', [t for t in traces if t['bbox'][1] < 400]),
                                       ('body', [t for t in traces if t['bbox'][1] >= 400])]:
                    if not selected:
                        ink[role] = 0
                        continue
                    box = fitz.Rect(selected[0]['bbox'])
                    for trace in selected[1:]:
                        box |= fitz.Rect(trace['bbox'])
                    pixels = raster(page, box + (-1, -1, 1, 1))
                    ink[role] = sum(min(pixels[j:j+3]) < 200 for j in range(0, len(pixels), 3))
                    assert ink[role] > 5, ('Invisible native text', role, i)
                if old:
                    assert raster(page) == raster(old[i]), 'Native notes raster changed'
                    assert traces == old[i].get_texttrace(), 'Native notes glyphs changed'
                    assert fonts(doc, page) == fonts(old, old[i]), 'Native font programs changed'
                pages.append({'label': expected[i]['label'], 'rasterSha256': sha(raster(page)),
                              'visibleInkPixels': ink,
                              'fontPrograms': fonts(doc, page),
                              'traces': [{'font': t['font'], 'size': t['size'], 'bbox': t['bbox'],
                                          'text': ''.join(chr(c[0]) for c in t['chars'])} for t in traces]})
        finally:
            if old:
                old.close()
    return pages


def inspect(directory):
    groups = []
    rejected = []
    for name in ['wide', 'portrait', 'square']:
        expected_file = directory / (name + '-expected.json')
        source = json.loads(expected_file.read_text())
        assert sha((directory / (name + '.pptx')).read_bytes()) == source['sha256']
        expected = [{'label': c['label'], 'text': ''.join(r['text'] for p in c['notes'] for r in p['runs']) if c.get('notes') else None} for c in source['controls']]
        # An explicit empty styled input has a body and an empty string.
        expected[3]['text'] = ''
        original = package(directory / (name + '.pptx'), expected)
        baseline = pdf(directory / (name + '-notes.pdf'), expected)
        rounds = []
        for number in range(1, 4):
            file = directory / f'{name}-round{number}.pptx'
            report = package(file, expected, original['geometry'])
            pdf_file = directory / f'{name}-round{number}-notes.pdf'
            pages = pdf(pdf_file, expected, directory / (name + '-notes.pdf'))
            rounds.append({'file': file.name, 'sha256': sha(file.read_bytes()),
                           'pdfSha256': sha(pdf_file.read_bytes()), 'package': report, 'pages': pages})
        wrong = expected.copy()
        wrong[0], wrong[1] = wrong[1], wrong[0]
        try:
            pdf(directory / (name + '-notes.pdf'), wrong)
        except AssertionError:
            rejected.append(name + ': swapped-copy-target')
        else:
            raise AssertionError('Accepted wrong copied slide')
        groups.append({'name': name, 'sourceSha256': source['sha256'], 'expectedSha256': sha(expected_file.read_bytes()),
                       'baselinePdfSha256': sha((directory / (name + '-notes.pdf')).read_bytes()),
                       'source': original, 'firstOpenPages': baseline, 'rounds': rounds})
    result = {'schemaVersion': 1, 'recordedAt': datetime.now(timezone.utc).isoformat(),
              'directory': str(directory), 'application': {'name': 'Microsoft PowerPoint for Mac', 'version': '16.104'},
              'groups': groups, 'rejectedControls': rejected,
              'total': {'publicAuthorInputs': 3, 'firstOpenPages': 12, 'nativeReopenPages': 36,
                        'nativeSaveCloseReopenRounds': 9, 'wrongCopyControlsRejected': 3},
              'knownFailure': 'PowerPoint drops the authored empty 12pt bold run at first save; per-page emptyStyledRunRetained records this failure explicitly. Geometry/raster evidence does not certify insertion styling.',
              'scope': 'Public fresh-author fixed notes layout only. Empty insertion-style native retention is failing. Imported artwork/fields, custom notes layouts, overflow pagination, native edit operations and cross-platform notes are not certified.'}
    Path('docs/testing/authored-notes-pages-evidence.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps(result['total']))


if __name__ == '__main__':
    inspect(Path(sys.argv[1]).resolve())
