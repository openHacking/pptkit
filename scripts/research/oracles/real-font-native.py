"""Independent named source slots and PDF font/glyph observations. Not native GUI certification."""
import sys, os, json, zipfile, xml.etree.ElementTree as E, hashlib
import fitz

expected_file, directory = sys.argv[1:]
expected = json.load(open(expected_file))
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
fonts = {f['family']: f['postscriptName'] for f in expected['fonts']}

def read_source(file, control):
    with zipfile.ZipFile(file) as z:
        root = E.fromstring(z.read('ppt/slides/slide1.xml'))
        shape = next(s for s in root.findall('.//' + P + 'sp') if s.find('.//' + P + 'cNvPr').get('name') == 'Probe')
        run = shape.find('.//' + A + 'r')
        assert run.find(A + 't').text == control['text']
        props = run.find(A + 'rPr')
        slots = [props.find(A + key) for key in ('latin', 'ea', 'cs')]
        assert [n.get('typeface') if n is not None else None for n in slots] == control['slots']
        return {'sourceRunCount': len(shape.findall('.//' + A + 'r')), 'slotValues': control['slots']}

def pdf(file):
    if not os.path.exists(file):
        return {'status': 'unavailable'}
    doc = fitz.open(file)
    assert len(doc) == 1
    page = doc[0]
    # Isolate body by its source frame, excluding ASCII title/font labels.
    spans = [s for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['bbox'][1] > 60]
    trace = [t for t in page.get_texttrace() if t['bbox'][1] > 60]
    glyphs = [{'unicode': c[0], 'glyph': c[1], 'origin': c[2]} for t in trace for c in t['chars']]
    return {'status': 'observed', 'fonts': sorted(set(s['font'] for s in spans)), 'extractedText': ''.join(s['text'] for s in spans), 'glyphs': glyphs, 'missingGlyphZeroCount': sum(g['glyph'] == 0 for g in glyphs), 'pdfSha256': hashlib.sha256(open(file, 'rb').read()).hexdigest()}

comparisons = []
for control in expected['controls']:
    name = control['id']
    source_file = os.path.join(directory, 'source', name + '.pptx')
    assert hashlib.sha256(open(source_file, 'rb').read()).hexdigest() == control['sourceSha256']
    slots = read_source(source_file, control)
    before = pdf(os.path.join(directory, 'pdf', name + '.pdf'))
    after = pdf(os.path.join(directory, 'native-pdf', name + '.pdf'))
    projected = sorted(set(s['family'] for s in control['projection']['selected']))
    expected_ps = sorted(set(fonts[family] for family in projected))
    actual = before.get('fonts', [])
    comparisons.append({'id': name, **slots, 'projectedFamilies': projected, 'sourcePdfFonts': actual, 'projectedFamiliesMatch': before['status'] == 'observed' and expected_ps == actual, 'sourcePdf': before, 'savedPdf': after})
for name in ('theme-explicit-sc', 'theme-empty-sc', 'theme-explicit-tc', 'theme-empty-arabic', 'direct-hebrew'):
    file = os.path.join(directory, 'pdf', name + '.pdf')
    if os.path.exists(file):
        book = fitz.open(file)
        book[0].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(os.path.join(directory, name + '.png'))
print(json.dumps({'sourceVerified': True, 'oracle': 'Python stdlib actual ZIP/XML slots and PyMuPDF actual body fonts/glyph traces before/after native save', 'comparisons': comparisons}))
