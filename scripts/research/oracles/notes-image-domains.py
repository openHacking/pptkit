"""Rect-only source repair across B and distinct per-notes font overrides."""
import importlib.util
import json
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as ET


def load(name, file):
    spec = importlib.util.spec_from_file_location(name, Path(__file__).with_name(file))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


pages = load('pages', 'native-notes-pages.py')
themes = load('themes', 'native-notes-themes.py')
NS, sha, payload, root = pages.NS, pages.sha, pages.payload, pages.root


def domains(file, controls):
    with zipfile.ZipFile(file) as z:
        rows = themes.domains(z, controls)
        for row in rows:
            master = pages.image_shape(themes.notes.root(z, row['masterPart']))
            geometry = master.find('p:spPr/a:prstGeom', NS)
            assert geometry is not None and geometry.get('prst') == 'rect', 'Missing master rect'
            pages.image_shape(themes.notes.root(z, row['notesPart']))
        return rows


def native_package(file, controls, fixed, source_domains):
    rows = domains(file, controls)
    assert rows == source_domains, 'Native notes/slide/theme identities changed'
    actual = payload(file)
    assert actual.keys() == fixed.keys()
    changed = [p for p in fixed if actual[p] != fixed[p]]
    assert set(changed) <= {'docProps/core.xml', 'docProps/app.xml'}, ('Unrelated native payload changed', changed)
    return rows, {part: {'beforeSha256': sha(fixed[part]), 'afterSha256': sha(actual[part]),
                         'beforeXml': fixed[part].decode(), 'afterXml': actual[part].decode()}
                  for part in changed}


def inspect(directory):
    groups = []
    for name in ['B', 'override']:
        folder = directory / name
        expected = json.loads((folder / 'expected.json').read_text())
        assert sha((folder / 'baseline.pptx').read_bytes()) == expected['sourceSha256']
        assert sha((folder / 'baseline-notes.pdf').read_bytes()) == expected['sourcePdfSha256']
        assert sha((folder / 'rect.pptx').read_bytes()) == expected['rectSha256']
        assert sha(Path(expected['manifest']).read_bytes()) == expected['manifestSha256']
        original, fixed = payload(folder / 'baseline.pptx'), payload(folder / 'rect.pptx')
        master = 'ppt/notesMasters/notesMaster1.xml'
        assert original.keys() == fixed.keys()
        assert [p for p in original if original[p] != fixed[p]] == [master], 'Changed non-master payload'
        old, new = root(original, master), root(fixed, master)
        props = pages.image_shape(new).find('p:spPr', NS)
        props.remove(props.find('a:prstGeom', NS))
        assert ET.tostring(old) == ET.tostring(new), 'Rect not the only source change'
        source_domains = domains(folder / 'rect.pptx', expected['controls'])
        if name == 'override':
            overridden = [r['overridePart'] for r in source_domains if r['variant'] == 'B']
            assert len(overridden) == len(set(overridden)) == 15 and all(overridden), 'Lost independent overrides'
            assert all(r['overridePart'] is None for r in source_domains if r['variant'] == 'A')
        else:
            assert all(r['overridePart'] is None for r in source_domains)
        controls = [{'label': 'notes-' + c['variant'] + '-' + c['id']} for c in expected['controls']]
        first = pages.check_pdf(folder / 'rect-notes.pdf', controls, folder / 'baseline-notes.pdf', False)
        rounds = []
        for number in range(1, 4):
            file = folder / f'round{number}.pptx'
            rows, metadata = native_package(file, expected['controls'], fixed, source_domains)
            pdf_file = folder / f'round{number}-notes.pdf'
            rendered = pages.check_pdf(pdf_file, controls, folder / 'rect-notes.pdf')
            rounds.append({'name': file.stem, 'sha256': sha(file.read_bytes()), 'pdfSha256': sha(pdf_file.read_bytes()),
                           'domains': rows, 'pages': rendered, 'changedMetadata': metadata})
        rejected = []
        try:
            pages.check_pdf(folder / 'baseline-notes.pdf', controls, folder / 'rect-notes.pdf')
        except AssertionError:
            rejected.append('actual-old-blank-image-pdf')
        else:
            raise AssertionError('Accepted old blank notes image')
        wrong = controls.copy()
        wrong[0], wrong[1] = wrong[1], wrong[0]
        try:
            pages.check_pdf(folder / 'round3-notes.pdf', wrong, folder / 'rect-notes.pdf')
        except AssertionError:
            rejected.append('wrong-page-map')
        else:
            raise AssertionError('Accepted wrong image page')
        negative_files = []
        output = folder / 'negative-controls'
        output.mkdir(exist_ok=True)
        mutations = [('master-rect', master, 'rect'),
                     ('slide-backlink', source_domains[0]['notesPart'].rsplit('/', 1)[0] + '/_rels/' + source_domains[0]['notesPart'].rsplit('/', 1)[1] + '.rels', 'backlink')]
        if name == 'override':
            target = next(r for r in source_domains if r['variant'] == 'B')
            rel = target['notesPart'].rsplit('/', 1)[0] + '/_rels/' + target['notesPart'].rsplit('/', 1)[1] + '.rels'
            mutations += [('missing-override', rel, 'remove-override'), ('override-font', target['overridePart'], 'font')]
        for label, part, kind in mutations:
            parts = payload(folder / 'round1.pptx')
            node = root(parts, part)
            if kind == 'rect':
                props = pages.image_shape(node).find('p:spPr', NS)
                props.remove(props.find('a:prstGeom', NS))
            elif kind == 'backlink':
                next(r for r in node if r.get('Type').endswith('/slide')).set('Target', '../slides/slide999.xml')
            elif kind == 'remove-override':
                node.remove(next(r for r in node if r.get('Type').endswith('/themeOverride')))
            else:
                latin = node.find('a:fontScheme/a:minorFont/a:latin', NS)
                assert latin is not None
                latin.set('typeface', 'Wrong domain font')
            parts[part] = ET.tostring(node, encoding='utf-8', xml_declaration=True)
            file = output / (label + '.pptx')
            with zipfile.ZipFile(file, 'w', zipfile.ZIP_DEFLATED) as z:
                for owner, body in parts.items():
                    z.writestr(owner, body)
            try:
                native_package(file, expected['controls'], fixed, source_domains)
            except (AssertionError, KeyError) as error:
                negative_files.append({'name': label, 'file': str(file), 'sha256': sha(file.read_bytes()), 'reason': str(error)})
            else:
                raise AssertionError('Accepted domain mutation: ' + label)
        groups.append({'name': name, 'directory': str(folder), 'expectedSha256': sha((folder / 'expected.json').read_bytes()),
                       'sourceSha256': expected['sourceSha256'], 'sourcePdfSha256': expected['sourcePdfSha256'],
                       'rectSha256': expected['rectSha256'], 'sourceDomains': source_domains,
                       'firstOpenPages': first, 'rounds': rounds, 'rejectedControls': rejected,
                       'actualPptxCorruptionsRejected': negative_files})
    result = {'schemaVersion': 1, 'recordedAt': datetime.now(timezone.utc).isoformat(), 'directory': str(directory),
              'application': {'name': 'Microsoft PowerPoint for Mac', 'version': '16.104'}, 'groups': groups,
              'total': {'sourcePageControls': 45, 'rectFirstOpenPages': 45, 'nativeReopenPages': 135,
                        'nativeSaveCloseReopenRounds': 6, 'blankImageAndWrongMapControlsRejected': 4,
                        'actualPptxCorruptionsRejected': sum(len(g['actualPptxCorruptionsRejected']) for g in groups)},
              'scope': 'B native-style baseline and one master with distinct per-notes overrides; rect-only repair, image identity and body/artwork retention. Original font-selection differences, shared-override repair admission and complete notes fields/artwork authoring remain separate.'}
    Path('docs/testing/notes-image-domains-evidence.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps(result['total']))


if __name__ == '__main__':
    inspect(Path(sys.argv[1]).resolve())
