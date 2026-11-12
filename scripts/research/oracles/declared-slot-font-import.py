"""Independent named v3 source-selection and scalar-preservation oracle."""
import sys, os, json, zipfile, xml.etree.ElementTree as E, re, copy, posixpath

expected_file, directory = sys.argv[1:]
snapshots = json.load(open(expected_file))
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
mutable = {'ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml'}
names = {'Hans': 'Simplified Source', 'Hant': 'Traditional Source', 'Jpan': 'Japanese Source', 'Hang': 'Korean Source', 'Arab': 'Arabic Source', 'Hebr': 'Hebrew Source', 'Grek': 'Greek Source'}

def package(file):
    with zipfile.ZipFile(file) as archive:
        members = [name for name in archive.namelist() if not name.endswith('/')]
        assert len(set(members)) == len(members)
        return {name: archive.read(name) for name in members}

def location(parts, source):
    node = E.fromstring(parts[source['part'].lstrip('/')])
    for index, piece in enumerate(source['path'].split('/@')[0].split('/')[1:]):
        match = re.fullmatch(r'([ap]):([^\[]+)\[(\d+)\]', piece)
        assert match, piece
        prefix, name, position = match.groups()
        tag = (A if prefix == 'a' else P) + name
        if index == 0:
            assert node.tag == tag and position == '1'
        else:
            children = node.findall(tag)
            assert len(children) >= int(position), 'missing actual source node'
            node = children[int(position) - 1]
    attribute = source['path'].split('/@')[1] if '/@' in source['path'] else None
    if attribute:
        assert attribute in node.attrib
    return node, attribute

def related(parts, source, kind):
    folder, name = posixpath.split(source)
    relations = E.fromstring(parts[posixpath.join(folder, '_rels', name + '.rels')])
    targets = [node.get('Target') for node in relations if node.get('Type') == R + kind and node.get('TargetMode') != 'External']
    assert len(targets) == 1
    return posixpath.normpath(posixpath.join(folder, targets[0]))

def rules(parts, snapshot):
    fonts = snapshot['fonts']
    assert len(fonts) == 42
    # Named fixture: no local or inherited source bucket exists anywhere in
    # presentation/drawing/notes style scopes. Theme generic slots still exist.
    for part, data in parts.items():
        if not part.startswith('ppt/') or not part.endswith('.xml') or part.startswith('ppt/theme/'):
            continue
        root = E.fromstring(data)
        assert root.find('.//' + A + 'ea') is None and root.find('.//' + A + 'cs') is None, 'unexpected actual native slot'
    mapped = blocked = 0
    for prop in fonts:
        node, attr = location(parts, prop['source'])
        code = prop['property'].split('.')[-1]
        notes = prop['target']['part'].startswith('/ppt/notesSlides/')
        expected = ('Notes ' if notes else '') + names[code]
        assert node.tag == A + 'font' and attr == 'typeface' and node.get('script') == code
        assert node.get('typeface') == expected, 'wrong actual supplemental source value'
        target_part = prop['target']['part'].lstrip('/')
        if notes:
            actual_theme = related(parts, related(parts, target_part, 'notesMaster'), 'theme')
        else:
            actual_theme = related(parts, related(parts, related(parts, target_part, 'slideLayout'), 'slideMaster'), 'theme')
        assert '/' + actual_theme == prop['source']['part'], 'wrong actual target theme chain'
        target, _ = location(parts, prop['target'])
        assert target.tag == A + 'r'
        assert target.find(A + 'rPr/' + A + 'latin').get('typeface') == '+mn-lt'
        assert any(dependency['part'] == prop['target']['part'] for dependency in prop['resolvedFrom'])
        for dependency in prop['resolvedFrom']:
            location(parts, dependency)
        if code == 'Grek':
            assert prop['status'] == 'mapped' and prop['value'] == expected
            mapped += 1
        else:
            assert prop['status'] == 'unsupported' and 'value' not in prop
            assert 'absent native slot' in prop['reason']
            blocked += 1
    assert (mapped, blocked) == (6, 36)
    return {'mapped': mapped, 'blocked': blocked}

def mask(data):
    return re.sub(rb'(<a:t(?:\s[^>]*)?>)[\s\S]*?(</a:t>)', rb'\1SCALAR\2', data)

def preserved(original, current):
    assert set(original) == set(current)
    changed = []
    for name in original:
        assert (mask(original[name]) == mask(current[name])) if name in mutable else (original[name] == current[name]), 'non-character source changed'
        if original[name] != current[name]:
            changed.append(name)
    return sorted(changed)

original = package(os.path.join(directory, snapshots[0]['file']))
rounds = []
for index, snapshot in enumerate(snapshots):
    parts = package(os.path.join(directory, snapshot['file']))
    count = rules(parts, snapshot)
    changed = preserved(original, parts)
    assert changed == ([] if index == 0 else sorted(mutable))
    expected_text = []
    def paragraphs(content):
        expected_text.extend(run['text'] for paragraph in content for run in paragraph['runs'])
    for element in snapshot['document']['slides'][0]['elements']:
        if element['type'] == 'text':
            paragraphs(element['content'])
        elif element['type'] == 'shape':
            paragraphs(element['text']['content'])
        elif element['type'] == 'table':
            for row in element['rows']:
                for cell in row['cells']:
                    paragraphs(cell['content'])
    paragraphs(snapshot['document']['slides'][0]['notes'])
    actual = [node.text or '' for part in ('ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml') for node in E.fromstring(parts[part]).findall('.//' + A + 't')]
    assert len(actual) == 6 and actual == expected_text
    controls = []
    for kind in ('mapped-value', 'blocked-value', 'missing-slot', 'theme-chain'):
        damaged = copy.deepcopy(parts)
        if kind in ('mapped-value', 'blocked-value'):
            code = 'Grek' if kind == 'mapped-value' else 'Hant'
            theme = 'ppt/theme/theme1.xml'
            before = ('script="' + code + '" typeface="' + names[code] + '"').encode()
            damaged[theme] = damaged[theme].replace(before, ('script="' + code + '" typeface="Wrong Source"').encode(), 1)
            assert damaged[theme] != parts[theme]
        elif kind == 'missing-slot':
            part = 'ppt/slides/slide1.xml'
            damaged[part] = damaged[part].replace(b'<a:latin typeface="+mn-lt"/>', b'<a:latin typeface="+mn-lt"/><a:ea typeface="Added East"/>', 1)
            assert damaged[part] != parts[part]
        else:
            part = 'ppt/slideMasters/_rels/slideMaster1.xml.rels'
            damaged[part] = damaged[part].replace(b'theme1.xml', b'theme2.xml')
            assert damaged[part] != parts[part]
        file = os.path.join(directory, 'negative-' + str(index) + '-' + kind + '.pptx')
        with zipfile.ZipFile(file, 'w') as archive:
            for name, data in damaged.items():
                archive.writestr(name, data)
        damaged = package(file)
        rejects = []
        for check in (lambda: rules(damaged, snapshot), lambda: preserved(original, damaged)):
            try:
                check()
            except AssertionError:
                rejects.append(True)
            else:
                rejects.append(False)
        assert all(rejects), 'actual ZIP negative control accepted'
        controls.append(kind)
    rounds.append({'round': index, **count, 'scalarTextsChecked': len(actual), 'changedParts': changed, 'negativeControlsRejectedByBothChecks': controls})
print(json.dumps({'sourceVerified': True, 'oracle': 'Python stdlib actual ZIP/XML, missing native scopes, actual target theme chains, source values/provenance, scalar text, exact non-character bytes and four actual ZIP negative controls', 'rounds': rounds}))
