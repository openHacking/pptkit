"""Independent authored family-to-XML checks; native output is observation only."""
import sys, os, json, zipfile, xml.etree.ElementTree as E, copy, fitz
expected_file, directory = sys.argv[1:]
expected = json.load(open(expected_file)); A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

def package(file):
    with zipfile.ZipFile(file) as z:
        return {n: z.read(n) for n in z.namelist() if not n.endswith('/')}

def font(value):
    return value if isinstance(value, str) else '+mj-lt' if value['theme'] == 'heading' else '+mn-lt'

def runs(content):
    return [r for p in content for r in p['runs']]

def element_runs(e):
    if e['type'] == 'text': return runs(e['content'])
    if e['type'] == 'shape': return runs(e['text']['content']) if e.get('text') else []
    if e['type'] == 'table': return [r for row in e['rows'] for c in row['cells'] for r in runs(c['content'])]
    if e['type'] == 'group': return [r for child in e['children'] for r in element_runs(child)]
    return []

def verify_xml(xml, wanted):
    root = E.fromstring(xml); actual = root.findall('.//' + A + 'r'); assert len(actual) == len(wanted), 'source run topology differs'
    for node, r in zip(actual, wanted):
        assert node.find(A + 't').text == r['text']
        style = r['style']; props = node.find(A + 'rPr')
        for tag, key in [('latin', 'fontFamily'), ('ea', 'eastAsianFontFamily'), ('cs', 'complexScriptFontFamily')]:
            assert props.find(A + tag).get('typeface') == font(style.get(key, style['fontFamily'])), (tag, style)
    return len(actual)

parts = package(os.path.join(directory, 'script-fonts.pptx'))
slide = 'ppt/slides/slide1.xml'; notes = 'ppt/notesSlides/notesSlide1.xml'
wanted = [r for e in expected['slides'][0]['elements'] for r in element_runs(e)]
checked = verify_xml(parts[slide], wanted) + verify_xml(parts[notes], runs(expected['slides'][0]['notes']))
damaged = copy.deepcopy(parts); root = E.fromstring(damaged[slide]); root.find('.//' + A + 'ea').set('typeface', 'Wrong Family'); damaged[slide] = E.tostring(root)
try:
    verify_xml(damaged[slide], wanted)
except AssertionError:
    rejected = True
else:
    rejected = False
assert rejected
native = {'status': 'unavailable'}
file = os.path.join(directory, 'native', 'script-fonts.pptx')
if os.path.exists(file):
    try:
        data = package(file); root = E.fromstring(data[slide]); actual = root.findall('.//' + A + 'r')
        native = {'status': 'observed', 'sourceRunCount': len(wanted), 'resavedRunCount': len(actual), 'runs': [{'text': n.find(A + 't').text, 'fonts': {tag: n.find(A + 'rPr').find(A + tag).get('typeface') if n.find(A + 'rPr') is not None and n.find(A + 'rPr').find(A + tag) is not None else None for tag in ('latin', 'ea', 'cs')}} for n in actual]}
    except (zipfile.BadZipFile, KeyError, E.ParseError) as e:
        native = {'status': 'unavailable', 'reason': str(e)}
pdf = []
for folder in ('source-pdf', 'saved-pdf'):
    file = os.path.join(directory, folder, 'script-fonts.pdf')
    if not os.path.exists(file):
        pdf.append({'kind': folder, 'status': 'unavailable'}); continue
    book = fitz.open(file); assert len(book) == 1; page = book[0]; spans = [s for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans']]
    png = folder + '.png'; page.get_pixmap(matrix=fitz.Matrix(2, 2)).save(os.path.join(directory, png))
    text = page.get_text()
    pdf.append({'kind': folder, 'status': 'observed', 'literalTextPresent': {s: s in text for s in ('Latin', '中文', 'العربية')}, 'text': text, 'spans': [{'text': s['text'], 'font': s['font'], 'size': s['size'], 'color': '%06X' % s['color']} for s in spans], 'png': png})
print(json.dumps({'sourceVerified': True, 'oracle': 'Python ZIP/ElementTree authored run font slots, deliberate selected ea mutation and native/PDF observations', 'sourceRunsChecked': checked, 'wrongFamilyRejected': rejected, 'native': native, 'pdf': pdf}))
