"""Independent declared-font/provenance and exact scalar-only source-save oracle."""
import sys, os, json, zipfile, xml.etree.ElementTree as E, re, copy, fitz
expected_file, directory = sys.argv[1:]; snapshots = json.load(open(expected_file))
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'; P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
mutable = {'ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml'}

def package(file):
    with zipfile.ZipFile(file) as z: return {n: z.read(n) for n in z.namelist() if not n.endswith('/')}

def location(parts, source):
    node = E.fromstring(parts[source['part'].lstrip('/')])
    for i, piece in enumerate(source['path'].split('/@')[0].split('/')[1:]):
        m = re.fullmatch(r'([^:]+):([^\[]+)\[(\d+)\]', piece); assert m, piece
        prefix, name, index = m.groups(); tag = (A if prefix == 'a' else P) + name
        if i == 0: assert node.tag == tag and index == '1'
        else:
            children = node.findall(tag); assert len(children) >= int(index), 'missing selected source'; node = children[int(index)-1]
    return node

def verify_fonts(parts, snapshot):
    theme = E.fromstring(parts['ppt/theme/theme1.xml']); count = 0
    for p in snapshot['fonts']:
        assert p['status'] == 'mapped', p
        node = location(parts, p['source']); value = node.get('typeface')
        assert value is not None
        if value.startswith('+'):
            m = re.fullmatch(r'\+(mj|mn)-(lt|ea|cs)', value); assert m
            group = theme.find('.//' + A + ('majorFont' if m[1] == 'mj' else 'minorFont'))
            value = group.find(A + ('latin' if m[2] == 'lt' else m[2])).get('typeface')
        assert value == p['value'], (value, p)
        count += 1
    return count

def masked(data):
    return re.sub(rb'(<a:t(?:\s[^>]*)?>)[\s\S]*?(</a:t>)', rb'\1SCALAR\2', data)

def verify_preserved(original, current):
    assert set(original) == set(current), 'part set changed'
    changed = []
    for part in original:
        if part in mutable:
            assert masked(original[part]) == masked(current[part]), 'non-character source bytes changed: ' + part
        else: assert original[part] == current[part], 'unmodified part changed: ' + part
        if original[part] != current[part]: changed.append(part)
    return sorted(changed)

original = package(os.path.join(directory, snapshots[0]['file'])); rounds = []
for i, snapshot in enumerate(snapshots):
    data = package(os.path.join(directory, snapshot['file'])); count = verify_fonts(data, snapshot); changed = verify_preserved(original, data)
    assert count == 18, count
    assert changed == ([] if i == 0 else sorted(mutable)), changed
    damaged = copy.deepcopy(data)
    damaged['ppt/slides/slide1.xml'] = re.sub(rb'(<a:ea\b[^>]*\btypeface=")[^"]*(")', rb'\1Wrong Family\2', damaged['ppt/slides/slide1.xml'], count=1)
    assert damaged['ppt/slides/slide1.xml'] != data['ppt/slides/slide1.xml']
    rejected = []
    for check in (lambda: verify_fonts(damaged, snapshot), lambda: verify_preserved(original, damaged)):
        try: check()
        except AssertionError: rejected.append(True)
        else: rejected.append(False)
    assert all(rejected), rejected
    # Scalar text itself must equal the approved reimport snapshot, including notes.
    expected_text = []
    def content(paragraphs): expected_text.extend(r['text'] for p in paragraphs for r in p['runs'])
    for e in snapshot['document']['slides'][0]['elements']:
        if e['type'] == 'text': content(e['content'])
        elif e['type'] == 'shape': content(e['text']['content'])
        elif e['type'] == 'table':
            for row in e['rows']:
                for c in row['cells']: content(c['content'])
    content(snapshot['document']['slides'][0]['notes'])
    actual_text = [n.text or '' for part in ('ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml') for n in E.fromstring(data[part]).findall('.//' + A + 't')]
    assert expected_text == actual_text
    rounds.append({'round': i, 'fontPropertiesChecked': count, 'changedParts': changed, 'scalarTextsChecked': len(actual_text), 'wrongFontRejectedByBothChecks': all(rejected)})

native = {'status': 'unavailable'}; file = os.path.join(directory, 'native', 'round-3.pptx')
if os.path.exists(file):
    try:
        data = package(file); root = E.fromstring(data['ppt/slides/slide1.xml'])
        native = {'status': 'observed', 'resavedRunCount': len(root.findall('.//' + A + 'r')), 'requestedFontNamesRetained': {name: name.encode() in data['ppt/slides/slide1.xml'] for name in ('Latin Face', 'Eastern Face', 'Complex Face')}}
    except (zipfile.BadZipFile, KeyError, E.ParseError) as e: native = {'status': 'unavailable', 'reason': str(e)}
pdf = []
for folder in ('source-pdf', 'saved-pdf'):
    file = os.path.join(directory, folder, 'round-3.pdf')
    if not os.path.exists(file): pdf.append({'kind': folder, 'status': 'unavailable'}); continue
    book = fitz.open(file); assert len(book) == 1; page = book[0]; text = page.get_text(); spans = [s for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans']]
    png = folder + '.png'; page.get_pixmap(matrix=fitz.Matrix(2, 2)).save(os.path.join(directory, png))
    pdf.append({'kind': folder, 'status': 'observed', 'scalarLabelPresent': {f'owner{i}': f'owner{i}' in text for i in range(1, 6)}, 'literalScriptTextPresent': {s: s in text for s in ('中文', 'العربية')}, 'text': text, 'fonts': sorted(set(s['font'] for s in spans)), 'png': png})
print(json.dumps({'sourceVerified': True, 'oracle': 'Python ZIP/ElementTree actual font declarations, selected mutation, exact bytes outside scalar text and native/PDF observations', 'rounds': rounds, 'native': native, 'pdf': pdf}))
