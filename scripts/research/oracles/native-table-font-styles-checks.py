"""Actual style ZIP/PDF mutations; no native baseline is overwritten."""
import importlib.util, json, re, sys, tempfile, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec = importlib.util.spec_from_file_location('table_style', Path(__file__).with_name('native-table-font-styles.py'))
o = importlib.util.module_from_spec(spec); spec.loader.exec_module(o)
p = Path(sys.argv[1]); expected = json.loads((p / 'expected.json').read_text()); rejected = []


def reject(name, fn):
    try: fn()
    except AssertionError: rejected.append(name); return
    raise AssertionError('Accepted corruption: ' + name)


def rewrite(destination, input, part, change):
    with zipfile.ZipFile(input) as source, zipfile.ZipFile(destination, 'w') as target:
        assert part in source.namelist()
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == part:
                root = ET.fromstring(data); change(root); data = ET.tostring(root)
            target.writestr(entry, data)


def paragraph(root): return root.find('.//a:tbl/a:tr/a:tc/a:txBody/a:p', o.NS)


with tempfile.TemporaryDirectory(prefix='pptkit-table-style-negative-') as temp:
    temp = Path(temp)
    for stage in o.STAGES:
        source = p / (stage + '-round1.pptx'); pdf = p / (stage + '-round1.pdf')
        verified = o.inspect_package(source, p, expected, stage)
        for row in verified['controls']:
            index, label, part = row['slide']-1, stage + '/' + row['id'], row['part']
            def flip_b(root): paragraph(root).findall('a:r', o.NS)[0].find('a:rPr', o.NS).set('b', '0' if stage == 'bold' else '1')
            def flip_i(root): paragraph(root).findall('a:r', o.NS)[1].find('a:rPr', o.NS).set('i', '0' if stage == 'italic' else '1')
            def insertion(root): paragraph(root).find('a:endParaRPr', o.NS).set('b', '0' if stage == 'bold' else '1')
            def local(root): ET.SubElement(paragraph(root).findall('a:r', o.NS)[0].find('a:rPr', o.NS), '{' + o.NS['a'] + '}ea', {'typeface': 'Wrong'})
            def text(root): paragraph(root).findall('a:r', o.NS)[1].find('a:t', o.NS).text = 'Wrong'
            for name, change in [('old-bold-flag', flip_b), ('new-italic-flag', flip_i), ('insertion-style', insertion), ('local-font', local), ('new-text', text)]:
                file = temp / 'changed.pptx'; rewrite(file, source, part, change)
                reject(label + '/' + name, lambda: o.inspect_package(file, p, expected, stage))
            with fitz.open(pdf) as doc:
                traces = [t for t in o.append.body_traces(doc[index]) if any(c[0] == 937 for c in t['chars'])]; assert len(traces) == 1
                b = traces[0]['bbox']; doc[index].draw_rect(fitz.Rect(b[0]-2,b[1]-2,b[2]+2,b[3]+2), color=None, fill=(1,1,1), overlay=True)
                file = temp / 'invisible.pdf'; doc.save(file)
            reject(label + '/extractable-invisible-omega', lambda: o.inspect_pdf(file, p, expected, stage))
            with fitz.open(pdf) as doc:
                fonts = [f for f in doc[index].get_fonts() if f[3].split('+')[-1] == traces[0]['font']]; assert fonts
                for font in fonts: doc.xref_set_key(font[0], 'BaseFont', '/Wrong')
                file = temp / 'resource.pdf'; doc.save(file)
            reject(label + '/font-resource', lambda: o.inspect_pdf(file, p, expected, stage))
            with fitz.open(pdf) as doc:
                changed = 0
                def shear(match):
                    global changed
                    if float(match.group(1)) != 200: return match.group(0)
                    changed += 1
                    values = list(match.groups()); values[2] = '11'
                    return 'BT ' + ' '.join(values[:6]) + ' Tm /' + values[6] + ' 1 Tf'
                for xref in doc[index].get_contents():
                    data = doc.xref_stream(xref).decode('latin1'); data = re.sub(o.MATRIX, shear, data); doc.update_stream(xref, data.encode('latin1'))
                assert changed == 2
                file = temp / 'shear.pdf'; doc.save(file)
            reject(label + '/text-matrix-shear', lambda: o.inspect_pdf(file, p, expected, stage))
        row = verified['controls'][0]
        file = temp / 'rule.pptx'
        def rule(root):
            if row['selector']['type'] == 'reference': node = root.find('a:themeElements/a:fontScheme/a:' + row['selector']['collection'] + 'Font', o.NS)
            else: node = next(n for n in root.findall('a:tblStyle', o.NS) if n.attrib['styleId'] == expected['controls'][row['slide']-1]['styleId']).find('a:wholeTbl/a:tcTxStyle/a:font', o.NS)
            node.find('a:font', o.NS).set('typeface', 'Wrong')
        rewrite(file, source, row['selector']['sourcePart'], rule)
        reject(stage + '/source-rule', lambda: o.inspect_package(file, p, expected, stage))
        with fitz.open(pdf) as doc:
            untouched = next(i for i in range(len(doc)) if i+1 not in expected['editTargets'])
            doc[untouched].draw_rect(fitz.Rect(500,180,530,200), color=None, fill=(0,0,0)); file = temp / 'neighbor.pdf'; doc.save(file)
        reject(stage + '/unmodified-neighbor', lambda: o.inspect_pdf(file, p, expected, stage))
        with fitz.open(pdf) as doc:
            doc.select([1,0]+list(range(2,len(doc)))); file = temp / 'order.pdf'; doc.save(file)
        reject(stage + '/page-order', lambda: o.inspect_pdf(file, p, expected, stage))
    # Removing a synthetic-bold draw mode must fail even though source text is still extractable.
    with fitz.open(p / 'bold-round1.pdf') as doc:
        index = next(i for i in expected['editTargets'] if o.named_face(expected['controls'][i-1], expected, 'bold')[1])-1
        changed = 0
        for xref in doc[index].get_contents():
            stream = doc.xref_stream(xref); changed += stream.count(b'2 Tr'); doc.update_stream(xref, stream.replace(b'2 Tr', b'0 Tr'))
        assert changed == 1; file = temp / 'no-stroke.pdf'; doc.save(file)
    reject('bold/synthetic-stroke-removed', lambda: o.inspect_pdf(file, p, expected, 'bold'))
print(json.dumps({'rejectedCount': len(rejected), 'rejectedMutations': rejected}))
