"""Independent stdlib ZIP/XML map/provenance and exact source preservation checks."""
import sys, os, json, zipfile, xml.etree.ElementTree as E, re, copy

expected_file, directory = sys.argv[1:]
snapshots = json.load(open(expected_file))
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
mutable = {'ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml'}

def package(file):
    with zipfile.ZipFile(file) as z:
        names = [n for n in z.namelist() if not n.endswith('/')]
        assert len(set(names)) == len(names), 'duplicate ZIP entries'
        return {n: z.read(n) for n in names}

def location(parts, source):
    node = E.fromstring(parts[source['part'].lstrip('/')])
    for i, piece in enumerate(source['path'].split('/@')[0].split('/')[1:]):
        match = re.fullmatch(r'([^:]+):([^\[]+)\[(\d+)\]', piece)
        assert match, piece
        prefix, name, index = match.groups()
        assert prefix in ('a', 'p'), prefix
        tag = (A if prefix == 'a' else P) + name
        if i == 0:
            assert node.tag == tag and index == '1'
        else:
            nodes = node.findall(tag)
            assert len(nodes) >= int(index), 'missing actual dependency/target'
            node = nodes[int(index)-1]
    attr = source['path'].split('/@')[1] if '/@' in source['path'] else None
    if attr:
        assert attr in node.attrib, 'missing actual attribute'
    return node, attr

def maps(parts, snapshot):
    assert len(snapshot['fonts']) == 29, len(snapshot['fonts'])
    for prop in snapshot['fonts']:
        assert prop['status'] == 'mapped'
        node, attr = location(parts, prop['source'])
        assert node.tag == A + 'font' and attr == 'typeface'
        assert node.get('script') == prop['property'].split('.')[-1]
        assert node.get('typeface') == prop['value'], 'wrong actual supplemental font'
        location(parts, prop['target'])
        assert any(d['part'] == '/ppt/slides/slide1.xml' for d in prop['resolvedFrom'])
        for dependency in prop['resolvedFrom']:
            location(parts, dependency)
    return len(snapshot['fonts'])

def mask(data):
    # Named generated fixture: scalar a:t leaves only. Not a general XML writer.
    return re.sub(rb'(<a:t(?:\s[^>]*)?>)[\s\S]*?(</a:t>)', rb'\1SCALAR\2', data)

def preserved(original, current):
    assert set(original) == set(current), 'part set changed'
    changed = []
    for name in original:
        if name in mutable:
            assert mask(original[name]) == mask(current[name]), 'non-character bytes changed'
        else:
            assert original[name] == current[name], 'unmodified payload changed'
        if original[name] != current[name]:
            changed.append(name)
    return sorted(changed)

original = package(os.path.join(directory, snapshots[0]['file']))
rounds = []
for index, snapshot in enumerate(snapshots):
    parts = package(os.path.join(directory, snapshot['file']))
    count = maps(parts, snapshot)
    changed = preserved(original, parts)
    assert changed == ([] if index == 0 else sorted(mutable))
    expected_text = []
    def content(paragraphs):
        expected_text.extend(r['text'] for p in paragraphs for r in p['runs'])
    for element in snapshot['document']['slides'][0]['elements']:
        if element['type'] == 'text':
            content(element['content'])
        elif element['type'] == 'shape':
            content(element['text']['content'])
        elif element['type'] == 'table':
            for row in element['rows']:
                for cell in row['cells']:
                    content(cell['content'])
    content(snapshot['document']['slides'][0]['notes'])
    actual = [n.text or '' for part in ('ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml') for n in E.fromstring(parts[part]).findall('.//' + A + 't')]
    assert len(actual) == 6 and actual == expected_text, 'actual scalar text differs'
    damaged = copy.deepcopy(parts)
    theme = 'ppt/theme/theme1.xml'
    damaged[theme] = damaged[theme].replace(b'script="Hant" typeface="Traditional Source"', b'script="Hant" typeface="Wrong Source"', 1)
    assert damaged[theme] != parts[theme]
    rejected = []
    for check in (lambda: maps(damaged, snapshot), lambda: preserved(original, damaged)):
        try:
            check()
        except AssertionError:
            rejected.append(True)
        else:
            rejected.append(False)
    assert all(rejected), 'negative control was accepted'
    rounds.append({'round': index, 'fontRulesChecked': count, 'scalarTextsChecked': len(actual), 'changedParts': changed, 'wrongMapRejectedByBothChecks': all(rejected)})
print(json.dumps({'sourceVerified': True, 'oracle': 'Python stdlib ZIP/ElementTree, actual map/provenance/target paths, scalar text, exact non-character source bytes and wrong-map negative controls', 'rounds': rounds}))
