"""Assemble named generated controls for manual native export; inspect actual PDF fonts.

This is deliberately limited to font-native-cases fixtures, not a PPTX merge API.
No font binaries are copied. Per-slide theme chains remain isolated.
"""
import sys, os, json, hashlib, zipfile, posixpath, copy, io, datetime, plistlib
import xml.etree.ElementTree as E
import fitz

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
OPC = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
for prefix, uri in [('p', P), ('a', A), ('r', R)]:
    E.register_namespace(prefix, uri)

def sha(data):
    return hashlib.sha256(data).hexdigest()

def xml(node):
    return E.tostring(node, encoding='utf-8', xml_declaration=True)

def related(archive, part, kind):
    folder, name = posixpath.split(part)
    relationships = E.fromstring(archive.read(posixpath.join(folder, '_rels', name + '.rels')))
    targets = [node.get('Target') for node in relationships if node.get('Type') == R + '/' + kind and node.get('TargetMode') != 'External']
    assert len(targets) == 1, (part, kind)
    return posixpath.normpath(posixpath.join(folder, targets[0]))

def assemble(directory, destination):
    expected = json.load(open(os.path.join(directory, 'expected.json')))
    parts, masters, slides = {}, E.Element('{' + P + '}sldMasterIdLst'), E.Element('{' + P + '}sldIdLst')
    rels = E.Element('{' + OPC + '}Relationships')
    types = E.Element('{' + CT + '}Types')
    E.SubElement(types, '{' + CT + '}Default', Extension='rels', ContentType='application/vnd.openxmlformats-package.relationships+xml')
    E.SubElement(types, '{' + CT + '}Default', Extension='xml', ContentType='application/xml')
    retained = []
    for index, control in enumerate(expected['controls']):
        source = open(os.path.join(directory, 'source', control['id'] + '.pptx'), 'rb').read()
        assert sha(source) == control['sourceSha256']
        with zipfile.ZipFile(io.BytesIO(source)) as archive:
            def renamed(name):
                return name.replace('1.xml', str(index + 1) + '.xml')
            if index == 0:
                presentation = E.fromstring(archive.read('ppt/presentation.xml'))
                for child in list(presentation):
                    if child.tag in ['{' + P + '}' + name for name in ('sldMasterIdLst', 'sldIdLst', 'notesMasterIdLst')]:
                        presentation.remove(child)
                parts['_rels/.rels'] = archive.read('_rels/.rels')
                for name in ('docProps/core.xml', 'docProps/app.xml'):
                    parts[name] = archive.read(name)
                for node in E.fromstring(archive.read('[Content_Types].xml')):
                    if node.get('PartName') in ('/ppt/presentation.xml', '/docProps/core.xml', '/docProps/app.xml'):
                        types.append(copy.deepcopy(node))
            admitted = [name for name in archive.namelist() if name.startswith(('ppt/slides/', 'ppt/slideLayouts/', 'ppt/slideMasters/', 'ppt/theme/'))]
            for name in admitted:
                data = archive.read(name)
                if name == 'ppt/slides/_rels/slide1.xml.rels':
                    slide_rels = E.fromstring(data)
                    for node in list(slide_rels):
                        if node.get('Type') == R + '/notesSlide':
                            slide_rels.remove(node)
                    data = xml(slide_rels)
                if name.endswith('.rels'):
                    relationships = E.fromstring(data)
                    for relation in relationships:
                        relation.set('Target', renamed(relation.get('Target')))
                    data = xml(relationships)
                if name == 'ppt/slideMasters/slideMaster1.xml':
                    master = E.fromstring(data)
                    master.find('.//{' + P + '}sldLayoutId').set('id', str(2147483649 + index * 2))
                    data = xml(master)
                if name != 'ppt/theme/theme2.xml':
                    parts[renamed(name)] = data
            for node in E.fromstring(archive.read('[Content_Types].xml')):
                if node.get('PartName', '').lstrip('/') in admitted and node.get('PartName') != '/ppt/theme/theme2.xml':
                    entry = copy.deepcopy(node)
                    entry.set('PartName', renamed(node.get('PartName')))
                    types.append(entry)
            for kind, target, listing, numeric in [
                ('slideMaster', 'slideMasters/slideMaster1.xml', masters, 2147483648 + index * 2),
                ('slide', 'slides/slide1.xml', slides, 256 + index),
            ]:
                identity = 'c' + str(index) + kind
                E.SubElement(rels, '{' + OPC + '}Relationship', Id=identity, Type=R + '/' + kind, Target=renamed(target))
                E.SubElement(listing, '{' + P + '}' + ('sldMasterId' if kind == 'slideMaster' else 'sldId'), {'id': str(numeric), '{' + R + '}id': identity})
            retained.append({'id': control['id'], 'slide': renamed('ppt/slides/slide1.xml'), 'theme': renamed('ppt/theme/theme1.xml'), 'slideSha256': sha(archive.read('ppt/slides/slide1.xml')), 'themeSha256': sha(archive.read('ppt/theme/theme1.xml'))})
    presentation.insert(0, masters)
    presentation.insert(1, slides)
    parts['ppt/presentation.xml'] = xml(presentation)
    parts['ppt/_rels/presentation.xml.rels'] = xml(rels)
    parts['[Content_Types].xml'] = xml(types)
    # Every internal relationship must still resolve inside this actual archive.
    for name, data in parts.items():
        if not name.endswith('.rels'):
            continue
        base = '' if name == '_rels/.rels' else posixpath.dirname(posixpath.dirname(name))
        for node in E.fromstring(data):
            assert node.get('TargetMode') != 'External'
            assert posixpath.normpath(posixpath.join(base, node.get('Target'))) in parts, (name, node.attrib)
    with zipfile.ZipFile(destination, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    manifest = {'sourceDirectory': directory, 'deckSha256': sha(open(destination, 'rb').read()), 'controls': retained}
    json.dump(manifest, open(destination + '.json', 'w'), indent=2)
    print(json.dumps({'deck': destination, 'pages': len(retained), 'sha256': manifest['deckSha256']}))

def inspect(directory, deck, pdf, destination):
    expected = json.load(open(os.path.join(directory, 'expected.json')))
    retained = json.load(open(deck + '.json'))
    assert len(retained['controls']) == len(expected['controls'])
    assert sha(open(deck, 'rb').read()) == retained['deckSha256']
    font_names = {font['family']: font['postscriptName'] for font in expected['fonts']}
    families = {font['id']: font['family'] for font in expected['fonts']}
    document = fitz.open(pdf)
    assert len(document) == len(expected['controls'])
    observations = []
    with zipfile.ZipFile(deck) as archive:
        presentation = E.fromstring(archive.read('ppt/presentation.xml'))
        presentation_relationships = {node.get('Id'): node for node in E.fromstring(archive.read('ppt/_rels/presentation.xml.rels'))}
        identities = list(presentation.find('{' + P + '}sldIdLst'))
        assert len(identities) == len(expected['controls'])
        for index, (control, retained_control) in enumerate(zip(expected['controls'], retained['controls'])):
            assert control['id'] == retained_control['id']
            relationship = presentation_relationships[identities[index].get('{' + R + '}id')]
            assert relationship.get('Type') == R + '/slide' and relationship.get('TargetMode') != 'External'
            actual_slide = posixpath.normpath(posixpath.join('ppt', relationship.get('Target')))
            assert actual_slide == retained_control['slide']
            for key in ('slide', 'theme'):
                assert sha(archive.read(retained_control[key])) == retained_control[key + 'Sha256']
            layout = related(archive, retained_control['slide'], 'slideLayout')
            master = related(archive, layout, 'slideMaster')
            theme_part = related(archive, master, 'theme')
            assert theme_part == retained_control['theme']
            slide = E.fromstring(archive.read(retained_control['slide']))
            probe = next(shape for shape in slide.findall('.//{' + P + '}sp') if shape.find('.//{' + P + '}cNvPr').get('name') == 'Probe')
            run = probe.find('.//{' + A + '}r')
            properties = run.find('{' + A + '}rPr')
            assert run.find('{' + A + '}t').text == control['text']
            assert properties.get('lang') == control['language']
            actual_slots = [properties.find('{' + A + '}' + slot) for slot in ('latin', 'ea', 'cs')]
            actual_slots = [node.get('typeface') if node is not None else None for node in actual_slots]
            assert actual_slots == control['slots']
            collection = (control.get('collection') or 'minor') + 'Font'
            theme = E.fromstring(archive.read(theme_part)).find('.//{' + A + '}' + collection)
            theme_slots = [theme.find('{' + A + '}' + slot).get('typeface') for slot in ('latin', 'ea', 'cs')]
            assert theme_slots == control.get('themeSlots', [families['latin'], '' if control.get('empty') else families['generic'], '' if control.get('empty') else families['generic']])
            supplemental = {script: families[family] for script, family in [('Hans', 'sc'), ('Hant', 'tc'), ('Jpan', 'ja'), ('Hang', 'ko'), ('Arab', 'ar'), ('Hebr', 'he'), ('Grek', 'greek')]}
            supplemental.update(control.get('supplemental', {}))
            supplemental = {script: family for script, family in supplemental.items() if family is not None}
            actual_supplemental = [(node.get('script'), node.get('typeface')) for node in theme.findall('{' + A + '}font')]
            assert len(dict(actual_supplemental)) == len(actual_supplemental)
            assert dict(actual_supplemental) == supplemental
            page = document[index]
            spans = [span for block in page.get_text('dict')['blocks'] if 'lines' in block for line in block['lines'] for span in line['spans']]
            assert any(control['id'] == span['text'] for span in spans), (index, control['id'])
            actual = sorted(set(span['font'] for span in spans if span['bbox'][1] > 60))
            projected = sorted(set(font_names[segment['family']] for segment in control['projection']['selected']))
            trace = [trace for trace in page.get_texttrace() if trace['bbox'][1] > 60]
            glyphs = [{'unicode': char[0], 'glyph': char[1], 'origin': char[2]} for trace in trace for char in trace['chars']]
            observations.append({'id': control['id'], 'page': index + 1, 'source': {'language': properties.get('lang'), 'slots': actual_slots, 'themePart': theme_part, 'collection': collection, 'themeSlots': theme_slots, 'supplemental': actual_supplemental}, 'actualFonts': actual, 'projectedFonts': projected, 'fontMatch': actual == projected, 'bodyText': ''.join(span['text'] for span in spans if span['bbox'][1] > 60), 'glyphs': glyphs, 'missingGlyphZeroCount': sum(char['glyph'] == 0 for char in glyphs)})
    result = {'status': 'observed', 'sourceDirectory': directory, 'deckPath': deck, 'pdfPath': pdf, 'recordedAt': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'deckSha256': retained['deckSha256'], 'pdfSha256': sha(open(pdf, 'rb').read()), 'sourceSlideAndThemePayloadsVerified': True, 'sourceControlValuesAndThemeChainsVerified': True, 'comparisons': observations}
    json.dump(result, open(destination, 'w'), ensure_ascii=False, indent=2)
    print(json.dumps({'pages': len(observations), 'mismatches': [row['id'] for row in observations if not row['fontMatch']]}))

def native_rounds(directory, source_deck, output_directory, destination):
    expected = json.load(open(os.path.join(directory, 'expected.json')))
    def inspect_declarations(file):
        with zipfile.ZipFile(file) as archive:
            def related(part, kind):
                folder, name = posixpath.split(part)
                relationships = E.fromstring(archive.read(posixpath.join(folder, '_rels', name + '.rels')))
                target = [node.get('Target') for node in relationships if node.get('Type') == R + '/' + kind]
                assert len(target) == 1, (part, kind)
                return posixpath.normpath(posixpath.join(folder, target[0]))
            presentation = E.fromstring(archive.read('ppt/presentation.xml'))
            relationships = {node.get('Id'): node.get('Target') for node in E.fromstring(archive.read('ppt/_rels/presentation.xml.rels'))}
            results = []
            for identity in presentation.find('{' + P + '}sldIdLst'):
                slide_part = posixpath.normpath(posixpath.join('ppt', relationships[identity.get('{' + R + '}id')]))
                slide = E.fromstring(archive.read(slide_part))
                probe = next(shape for shape in slide.findall('.//{' + P + '}sp') if shape.find('.//{' + P + '}cNvPr').get('name') == 'Probe')
                layout = related(slide_part, 'slideLayout')
                master = related(layout, 'slideMaster')
                theme = related(master, 'theme')
                minor = E.fromstring(archive.read(theme)).find('.//{' + A + '}minorFont')
                runs = probe.findall('.//{' + A + '}r')
                def slots(node):
                    return [(child.tag.split('}')[-1], child.get('typeface')) for child in node if child.tag in ['{' + A + '}' + name for name in ('latin', 'ea', 'cs')]] if node is not None else []
                results.append({'slidePart': slide_part, 'layoutPart': layout, 'masterPart': master, 'themePart': theme, 'runs': [{'text': run.find('{' + A + '}t').text, 'language': run.find('{' + A + '}rPr').get('lang'), 'slots': slots(run.find('{' + A + '}rPr'))} for run in runs], 'localListFonts': [slots(node) for node in probe.findall('.//{' + A + '}lstStyle//{' + A + '}defRPr')], 'themeGeneric': slots(minor), 'themeSupplemental': [(node.get('script'), node.get('typeface')) for node in minor.findall('{' + A + '}font')]})
            assert len(results) == len(expected['controls'])
            return results
    source = inspect_declarations(source_deck)
    initial_pdf = os.path.splitext(source_deck)[0] + '.pdf'
    initial = fitz.open(initial_pdf)
    def pdf_rows(file):
        document = fitz.open(file)
        assert len(document) == len(expected['controls'])
        rows = []
        for index, control in enumerate(expected['controls']):
            page = document[index]
            spans = [span for block in page.get_text('dict')['blocks'] if 'lines' in block for line in block['lines'] for span in line['spans']]
            assert any(span['text'] == control['id'] for span in spans)
            body = [span for span in spans if span['bbox'][1] > 60]
            rows.append({'id': control['id'], 'fonts': sorted(set(span['font'] for span in body)), 'text': ''.join(span['text'] for span in body), 'rasterMatchesInitial': page.get_pixmap().samples == initial[index].get_pixmap().samples})
        return rows
    stages = [{'name': 'source', 'pptxSha256': sha(open(source_deck, 'rb').read()), 'pdfSha256': sha(open(initial_pdf, 'rb').read()), 'declarations': source, 'pdf': pdf_rows(initial_pdf)}]
    semantic_keys = ['runs', 'localListFonts', 'themeGeneric', 'themeSupplemental']
    for round_index in range(1, 4):
        stem = os.path.join(output_directory, 'font-controls-round' + str(round_index))
        declarations = inspect_declarations(stem + '.pptx')
        changes = [{'id': control['id'], 'changedFields': [key for key in semantic_keys if before[key] != after[key]]} for control, before, after in zip(expected['controls'], source, declarations)]
        stages.append({'name': 'round' + str(round_index), 'pptxSha256': sha(open(stem + '.pptx', 'rb').read()), 'pdfSha256': sha(open(stem + '.pdf', 'rb').read()), 'declarations': declarations, 'declarationChanges': [row for row in changes if row['changedFields']], 'pdf': pdf_rows(stem + '.pdf')})
    # Mutate an actual package, then use the same ZIP/relationship/declaration reader.
    wrong_file = os.path.join(output_directory, 'wrong-font-control.pptx')
    with zipfile.ZipFile(source_deck) as archive, zipfile.ZipFile(wrong_file, 'w') as output:
        for name in archive.namelist():
            data = archive.read(name)
            if name == source[0]['slidePart']:
                root = E.fromstring(data)
                probe = next(shape for shape in root.findall('.//{' + P + '}sp') if shape.find('.//{' + P + '}cNvPr').get('name') == 'Probe')
                probe.find('.//{' + A + '}rPr/{' + A + '}ea').set('typeface', 'Wrong Font')
                data = xml(root)
            output.writestr(name, data)
    corrupt = inspect_declarations(wrong_file)
    assert source[0]['runs'] != corrupt[0]['runs']
    assert all(source[index] == corrupt[index] for index in range(1, len(source)))
    with open('/Applications/Microsoft PowerPoint.app/Contents/Info.plist', 'rb') as file:
        application = plistlib.load(file)
    result = {'status': 'observed', 'recordedAt': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'application': 'Microsoft PowerPoint for Mac ' + application['CFBundleShortVersionString'], 'applicationBuild': application['CFBundleVersion'], 'sourceDirectory': directory, 'artifactsDirectory': output_directory, 'stages': stages, 'wrongFontControlRejected': True, 'limitations': ['Named font declarations and raster equality only', 'No text/style/theme editing certification', 'No complete package byte preservation claim']}
    json.dump(result, open(destination, 'w'), ensure_ascii=False, indent=2)
    print(json.dumps({'stages': [{'name': stage['name'], 'declarationChanges': stage.get('declarationChanges', []), 'rasterMismatches': [row['id'] for row in stage['pdf'] if not row['rasterMatchesInitial']]} for stage in stages]}))

if __name__ == '__main__':
    mode, *arguments = sys.argv[1:]
    if mode == 'assemble':
        assemble(*arguments)
    elif mode == 'inspect':
        inspect(*arguments)
    elif mode == 'native-rounds':
        native_rounds(*arguments)
    else:
        raise ValueError('Expected assemble SOURCE_DIR DECK, inspect SOURCE_DIR DECK PDF RESULT or native-rounds SOURCE_DIR DECK OUTPUT_DIR RESULT')
