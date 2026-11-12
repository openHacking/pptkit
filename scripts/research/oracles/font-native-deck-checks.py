"""Optional negative controls against actual native-research artifacts; no font install.

Usage: python3 .../font-native-deck-checks.py SOURCE_DIR DECK PDF
The mutations update transport hashes deliberately, so expected XML values and
actual relationship chains must detect them rather than a hash-only guard.
"""
import sys, os, json, zipfile, tempfile, copy, importlib.util
import xml.etree.ElementTree as E

specification = importlib.util.spec_from_file_location('font_native_deck', os.path.join(os.path.dirname(__file__), 'font-native-deck.py'))
oracle = importlib.util.module_from_spec(specification)
specification.loader.exec_module(oracle)
source_directory, source_deck, source_pdf = sys.argv[1:]
manifest = json.load(open(source_deck + '.json'))
with zipfile.ZipFile(source_deck) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}

with tempfile.TemporaryDirectory(prefix='font-control-checks-', dir='.tmp') as temporary:
    oracle.inspect(source_directory, source_deck, source_pdf, os.path.join(temporary, 'positive.json'))
    positive = json.load(open(os.path.join(temporary, 'positive.json')))
    expected = json.load(open(os.path.join(source_directory, 'expected.json')))
    assert len(positive['comparisons']) == len(expected['controls'])
    assert [row['id'] for row in positive['comparisons']] == [row['id'] for row in expected['controls']]
    checked = []
    for kind in ('wrong-slot', 'wrong-language', 'wrong-supplemental', 'wrong-theme-chain', 'wrong-page-order'):
        parts, metadata = copy.deepcopy(original), copy.deepcopy(manifest)
        retained = metadata['controls'][0]
        if kind in ('wrong-slot', 'wrong-language'):
            root = E.fromstring(parts[retained['slide']])
            probe = next(shape for shape in root.findall('.//{' + oracle.P + '}sp') if shape.find('.//{' + oracle.P + '}cNvPr').get('name') == 'Probe')
            properties = probe.find('.//{' + oracle.A + '}rPr')
            if kind == 'wrong-slot':
                properties.find('{' + oracle.A + '}ea').set('typeface', 'Wrong Real Font')
            else:
                properties.set('lang', 'ja-JP')
            parts[retained['slide']] = oracle.xml(root)
        elif kind == 'wrong-supplemental':
            root = E.fromstring(parts[retained['theme']])
            root.find('.//{' + oracle.A + '}minorFont/{' + oracle.A + '}font').set('typeface', 'Wrong Supplemental Font')
            parts[retained['theme']] = oracle.xml(root)
        elif kind == 'wrong-theme-chain':
            relation_part = 'ppt/slides/_rels/slide1.xml.rels'
            root = E.fromstring(parts[relation_part])
            next(node for node in root if node.get('Type') == oracle.R + '/slideLayout').set('Target', '../slideLayouts/slideLayout2.xml')
            parts[relation_part] = oracle.xml(root)
        else:
            root = E.fromstring(parts['ppt/presentation.xml'])
            slides = root.find('{' + oracle.P + '}sldIdLst')
            second = slides[1]
            slides.remove(second)
            slides.insert(0, second)
            parts['ppt/presentation.xml'] = oracle.xml(root)
        deck = os.path.join(temporary, kind + '.pptx')
        with zipfile.ZipFile(deck, 'w') as output:
            for name, data in parts.items():
                output.writestr(name, data)
        metadata['deckSha256'] = oracle.sha(open(deck, 'rb').read())
        for item in metadata['controls']:
            for field in ('slide', 'theme'):
                item[field + 'Sha256'] = oracle.sha(parts[item[field]])
        json.dump(metadata, open(deck + '.json', 'w'))
        try:
            oracle.inspect(source_directory, deck, source_pdf, os.path.join(temporary, 'invalid.json'))
        except AssertionError:
            checked.append(kind)
        else:
            raise AssertionError('Oracle accepted mutated actual control: ' + kind)
    print(json.dumps({'positivePages': len(positive['comparisons']), 'negativeControlsRejected': checked}))
