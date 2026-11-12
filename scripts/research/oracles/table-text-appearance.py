"""Independent source location/color/typeface checks and native PDF observations."""
import sys, os, json, zipfile, xml.etree.ElementTree as E, re, copy, fitz
expected_file, directory = sys.argv[1:]
expected = json.load(open(expected_file))
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'

def package(file):
    with zipfile.ZipFile(file) as z:
        return {n: z.read(n) for n in z.namelist() if not n.endswith('/')}

def location(parts, source):
    node = E.fromstring(parts[source['part'].lstrip('/')])
    path = source['path'].split('/@')[0]
    for i, piece in enumerate(path.split('/')[1:]):
        match = re.fullmatch(r'([^:]+):([^\[]+)\[(\d+)\]', piece)
        assert match, piece
        prefix, name, index = match.groups(); tag = (A if prefix == 'a' else P) + name
        if i == 0:
            assert node.tag == tag and index == '1'
        else:
            matches = node.findall(tag); assert len(matches) >= int(index), 'missing source declaration'
            node = matches[int(index) - 1]
    return node

def verify(parts, scene):
    theme = E.fromstring(parts['ppt/theme/theme1.xml'])
    colors = {n.tag.split('}')[-1]: list(n)[0].get('val') for n in theme.find('.//' + A + 'clrScheme')}
    checked = 0
    for p in scene['provenance']['properties']:
        if p['status'] != 'mapped' or p['property'] not in ('run.color', 'run.fontFamily'):
            continue
        node = location(parts, p['source'])
        if p['property'] == 'run.color':
            value = colors.get(node.get('val')) if node.tag == A + 'schemeClr' else node.get('val')
            assert value == p['value'], (p, value)
        else:
            if node.tag == A + 'fontRef':
                name = 'minorFont' if node.get('idx') == 'minor' else 'majorFont'
                assert theme.find('.//' + A + name).find(A + 'latin').get('typeface') == p['value']
            else:
                assert node.tag == A + 'latin'
                assert node.get('typeface') in (p['value'], '+mn-lt')
                for dep in p.get('resolvedFrom', []):
                    assert location(parts, dep).get('typeface') in (p['value'], '+mn-ea', '+mn-cs')
        checked += 1
    return checked

scenes = []
for scene in expected:
    name = scene['name']; source = package(os.path.join(directory, name + '.pptx')); checked = verify(source, scene)
    bad = copy.deepcopy(source)
    # Corrupt the effective source color itself, not an unrelated lookalike tree.
    p = next(p for p in scene['provenance']['properties'] if p['property'] == 'run.color' and p['status'] == 'mapped')
    part = p['source']['part'].lstrip('/'); root = E.fromstring(bad[part])
    # Replace the exact indexed source node using the original lexical path walk.
    path = p['source']['path'].split('/@')[0].split('/')[2:]; target = root
    for piece in path[:-1]:
        prefix, local, index = re.fullmatch(r'([^:]+):([^\[]+)\[(\d+)\]', piece).groups()
        target = target.findall((A if prefix == 'a' else P) + local)[int(index)-1]
    prefix, local, index = re.fullmatch(r'([^:]+):([^\[]+)\[(\d+)\]', path[-1]).groups()
    old = target.findall((A if prefix == 'a' else P) + local)[int(index)-1]
    old.set('val', 'FF00FF')
    bad[part] = E.tostring(root)
    try:
        verify(bad, scene)
    except AssertionError:
        rejected = True
    else:
        rejected = False
    assert rejected, 'oracle did not reject selected source color mutation'
    pdf = []
    for folder in ('source-pdf', 'saved-pdf'):
        file = os.path.join(directory, folder, name + '.pdf')
        if not os.path.exists(file):
            pdf.append({'kind': folder, 'status': 'unavailable'}); continue
        book = fitz.open(file); assert len(book) == 1; page = book[0]
        spans = [span for block in page.get_text('dict')['blocks'] if 'lines' in block for line in block['lines'] for span in line['spans']]
        differences, observed = [], []
        for r, row in enumerate(scene['table']['rows']):
            for c, cell in enumerate(row['cells']):
                label = f'R{r}C{c}'; style = cell['content'][0]['runs'][0]['style']
                matches = [s for s in spans if label in s['text']]
                if len(matches) != 1:
                    differences.append({'label': label, 'issue': 'missing or ambiguous PDF text'}); continue
                span = matches[0]; value = '%06X' % span['color']
                observed.append({'label': label, 'color': value, 'font': span['font'], 'size': span['size']})
                if value != style['color']:
                    differences.append({'label': label, 'property': 'color', 'expected': style['color'], 'actual': value})
        png = name + '-' + folder + '.png'; page.get_pixmap(matrix=fitz.Matrix(2, 2)).save(os.path.join(directory, png))
        pdf.append({'kind': folder, 'status': 'pass' if not differences else 'partial', 'differences': differences, 'observed': observed, 'png': png})
    scenes.append({'name': name, 'independentSourceProperties': checked, 'selectedColorMutationRejected': rejected, 'pdf': pdf})
print(json.dumps({'verified': True, 'oracle': 'Python ZIP/ElementTree source color/typeface dependencies, selected color mutation, independent PDF spans', 'scenes': scenes}))
