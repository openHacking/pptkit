"""Independent stdlib OPC/XML notes style/font/color/type-selection and byte oracle."""
import sys, os, json, re, copy, posixpath, zipfile, xml.etree.ElementTree as E

expected_file, directory = sys.argv[1:]
snapshots = json.load(open(expected_file))
mutable = {'deck/z.xml', 'deck/notes/z.xml', 'deck/notes/a.xml'}

def package(file):
    with zipfile.ZipFile(file) as z:
        names = [n for n in z.namelist() if not n.endswith('/')]
        assert len(set(names)) == len(names)
        return {n: z.read(n) for n in names}

def ns(tag):
    return tag.split('}')[0] + '}'

def location(parts, source):
    node = E.fromstring(parts[source['part'].lstrip('/')])
    p = ns(E.fromstring(parts['deck/notes/z.xml']).tag)
    a = ns(E.fromstring(parts['deck/notes-theme.xml']).tag)
    for index, piece in enumerate(source['path'].split('/@')[0].split('/')[1:]):
        m = re.fullmatch(r'([^:]+):([^\[]+)\[(\d+)\]', piece)
        assert m, piece
        prefix, name, ordinal = m.groups()
        assert prefix in ('a', 'p')
        tag = (a if prefix == 'a' else p) + name
        if index == 0:
            assert node.tag == tag and ordinal == '1'
        else:
            nodes = node.findall(tag)
            assert len(nodes) >= int(ordinal), 'missing actual source/dependency/target'
            node = nodes[int(ordinal)-1]
    attr = source['path'].split('/@')[1] if '/@' in source['path'] else None
    if attr:
        assert attr in node.attrib
    return node, attr

def related(parts, owner, role):
    relpart = posixpath.join(posixpath.dirname(owner), '_rels', posixpath.basename(owner) + '.rels')
    entries = [r for r in E.fromstring(parts[relpart]) if r.get('Type', '').endswith('/' + role)]
    assert len(entries) == 1 and entries[0].get('TargetMode') != 'External'
    part = posixpath.normpath(posixpath.join(posixpath.dirname(owner), entries[0].get('Target')))
    assert part in parts
    return part

def styles(parts, snapshot):
    assert len(snapshot['properties']) == 22
    a = ns(E.fromstring(parts['deck/notes-theme.xml']).tag)
    p = ns(E.fromstring(parts['deck/notes/z.xml']).tag)
    for prop in snapshot['properties']:
        assert prop['status'] == 'mapped'
        owner = prop['target']['part'].lstrip('/')
        master_part = related(parts, owner, 'notesMaster')
        theme_part = related(parts, master_part, 'theme')
        theme = E.fromstring(parts[theme_part])
        node, attr = location(parts, prop['source'])
        location(parts, prop['target'])
        for dep in prop.get('resolvedFrom', []):
            location(parts, dep)
        key, value = prop['property'], prop['value']
        if key.startswith('run.scriptFontFamilies.'):
            assert prop['source']['part'].lstrip('/') == theme_part
            assert node.get('script') == key.split('.')[-1] and node.get('typeface') == value
        elif key in ('run.fontFamily', 'run.eastAsianFontFamily', 'run.complexScriptFontFamily'):
            token = node.get('typeface')
            match = re.fullmatch(r'\+(mj|mn)-(lt|ea|cs)', token)
            assert match
            group = theme.find('.//' + a + ('majorFont' if match[1] == 'mj' else 'minorFont'))
            family = group.find(a + ('latin' if match[2] == 'lt' else match[2])).get('typeface')
            assert family == value
        elif key == 'run.fontSize':
            assert float(node.get('sz')) / 100 == value == 22
            assert prop['level'] == 'notes-master-placeholder'
            types = [location(parts, dep)[0] for dep in prop['resolvedFrom']]
            assert len(types) == 2 and all(n.get('type') == 'body' for n in types)
            assert sorted(n.get('idx') for n in types) == ['1', '99'], 'must prove type rather than idx match'
        elif key == 'run.italic':
            assert node.get('i') == '1' and value is True
        elif key == 'run.language':
            assert node.get('lang') == value == 'zh-TW'
        elif key in ('run.color', 'run.colorOpacity'):
            master = E.fromstring(parts[master_part])
            slot = master.find(p + 'clrMap').get(node.get('val'))
            rgb = theme.find('.//' + a + slot).find(a + 'srgbClr').get('val')
            assert value == (rgb if key == 'run.color' else 1)
        else:
            raise AssertionError('unverified property ' + key)
    return len(snapshot['properties'])

def mask(data):
    return re.sub(rb'(<a:t(?:\s[^>]*)?>)[\s\S]*?(</a:t>)', rb'\1SCALAR\2', data)

def preserved(original, current):
    assert set(original) == set(current)
    changed = []
    for part in original:
        if part in mutable:
            assert mask(original[part]) == mask(current[part]), 'non-character bytes changed'
        else:
            assert original[part] == current[part], 'immutable payload changed'
        if original[part] != current[part]:
            changed.append(part)
    return sorted(changed)

original = package(os.path.join(directory, snapshots[0]['file']))
rounds = []
for index, snapshot in enumerate(snapshots):
    parts = package(os.path.join(directory, snapshot['file']))
    count = styles(parts, snapshot)
    changed = preserved(original, parts)
    assert changed == ([] if index == 0 else sorted(mutable))
    for page, letter in enumerate(('z', 'a')):
        content = E.fromstring(parts[f'deck/notes/{letter}.xml'])
        a = ns(E.fromstring(parts['deck/notes-theme.xml']).tag)
        actual = [n.text or '' for n in content.findall('.//' + a + 't')]
        approved = [r['text'] for p in snapshot['document']['slides'][page]['notes'] for r in p['runs']]
        assert actual == approved and len(actual) == 1
    controls = []
    for part, before, after in [('deck/notes-theme.xml', b'typeface="Notes Traditional"', b'typeface="Wrong Notes"'), ('deck/notes-master.xml', b'type="body"', b'type="ftr"')]:
        damaged = copy.deepcopy(parts)
        damaged[part] = damaged[part].replace(before, after, 1)
        assert damaged[part] != parts[part]
        rejected = []
        for check in (lambda: styles(damaged, snapshot), lambda: preserved(original, damaged)):
            try:
                check()
            except AssertionError:
                rejected.append(True)
            else:
                rejected.append(False)
        assert all(rejected), 'negative control accepted'
        controls.append({'part': part, 'rejectedByBothChecks': True})
    rounds.append({'round': index, 'stylePropertiesChecked': count, 'notesTextsChecked': 2, 'changedParts': changed, 'negativeControls': controls})
print(json.dumps({'sourceVerified': True, 'oracle': 'Python stdlib OPC/XML actual notes-master/theme relationships, type matching, style/font/color values and non-character bytes', 'rounds': rounds}))
