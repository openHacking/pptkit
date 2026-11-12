"""Actual ZIP/PDF corruption controls for the native table append oracle."""
import importlib.util, json, sys, tempfile, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec = importlib.util.spec_from_file_location('table_edits', Path(__file__).with_name('native-table-font-edits.py'))
o = importlib.util.module_from_spec(spec)
spec.loader.exec_module(o)
p = Path(sys.argv[1]); expected = json.loads((p / 'expected.json').read_text())
accepted = p / 'append-right-round1.pptx'
verified = o.inspect_package(accepted, p, expected)
rejected = []


def reject(name, fn):
    try: fn()
    except AssertionError: rejected.append(name); return
    raise AssertionError('Accepted corruption: ' + name)


def rewrite(destination, part, change):
    with zipfile.ZipFile(accepted) as source, zipfile.ZipFile(destination, 'w') as target:
        assert part in source.namelist()
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == part:
                root = ET.fromstring(data); change(root); data = ET.tostring(root)
            target.writestr(entry, data)


def paragraph(root): return root.find('.//a:tbl/a:tr/a:tc/a:txBody/a:p', o.NS)


with tempfile.TemporaryDirectory(prefix='pptkit-table-append-negative-') as temp:
    temp = Path(temp)
    for row in verified['controls']:
        part, index, label = row['part'], row['slide'] - 1, row['id']
        def reorder(root):
            para = paragraph(root); runs = para.findall('a:r', o.NS)
            para.remove(runs[1]); para.insert(list(para).index(runs[0]), runs[1])
        def local(root): ET.SubElement(paragraph(root).findall('a:r', o.NS)[1].find('a:rPr', o.NS), '{' + o.NS['a'] + '}latin', {'typeface': 'Wrong'})
        changes = [('prepend', reorder), ('new-language', lambda r: paragraph(r).findall('a:r', o.NS)[1].find('a:rPr', o.NS).set('lang', 'ko-KR')), ('local-font', local), ('old-text', lambda r: setattr(paragraph(r).findall('a:r', o.NS)[0].find('a:t', o.NS), 'text', 'Wrong')), ('style', lambda r: setattr(r.find('.//a:tbl/a:tblPr/a:tableStyleId', o.NS), 'text', 'Wrong'))]
        for name, change in changes:
            file = temp / (label + '-' + name + '.pptx'); rewrite(file, part, change)
            reject(label + '/' + name, lambda: o.inspect_package(file, p, expected))
        with fitz.open(p / 'append-right-round1.pdf') as doc:
            omega = [t for t in o.body_traces(doc[index]) if any(c[0] == 937 for c in t['chars'])]
            assert len(omega) == 1
            b = omega[0]['bbox']; doc[index].draw_rect(fitz.Rect(b[0]-2, b[1]-2, b[2]+2, b[3]+2), color=None, fill=(1,1,1), overlay=True)
            file = temp / (label + '-invisible.pdf'); doc.save(file)
        reject(label + '/extractable-invisible-omega', lambda: o.inspect_pdf(file, p, expected))
        with fitz.open(p / 'append-right-round1.pdf') as doc:
            resources = [f for f in doc[index].get_fonts() if f[3].split('+')[-1] == 'Georgia']; assert resources
            for font in resources: doc.xref_set_key(font[0], 'BaseFont', '/Wrong')
            file = temp / (label + '-resource.pdf'); doc.save(file)
        reject(label + '/actual-font-resource', lambda: o.inspect_pdf(file, p, expected))
    file = temp / 'rule.pptx'
    row = verified['controls'][0]
    def rule(root):
        if row['selector']['type'] == 'reference': node = root.find('a:themeElements/a:fontScheme/a:' + row['selector']['collection'] + 'Font', o.NS)
        else: node = next(n for n in root.findall('a:tblStyle', o.NS) if n.attrib['styleId'] == expected['controls'][row['slide']-1]['styleId']).find('a:wholeTbl/a:tcTxStyle/a:font', o.NS)
        node.find('a:font', o.NS).set('typeface', 'Wrong')
    rewrite(file, row['selector']['sourcePart'], rule)
    reject('source-script-rule', lambda: o.inspect_package(file, p, expected))
    with fitz.open(p / 'append-right-round1.pdf') as doc:
        doc.select([1, 0] + list(range(2, len(doc)))); file = temp / 'order.pdf'; doc.save(file)
    reject('PDF-order', lambda: o.inspect_pdf(file, p, expected))
    with fitz.open(p / 'append-right-round1.pdf') as doc:
        untouched = next(i for i in range(len(doc)) if i+1 not in expected['editTargets'])
        doc[untouched].draw_rect(fitz.Rect(500, 180, 530, 200), color=None, fill=(0,0,0)); file = temp / 'neighbor.pdf'; doc.save(file)
    reject('unmodified-neighbor-raster', lambda: o.inspect_pdf(file, p, expected))
    if (p / 'append-round1.pptx').exists(): reject('actual-native-prepend-trial', lambda: o.inspect_package(p / 'append-round1.pptx', p, expected))
print(json.dumps({'rejectedCount': len(rejected), 'rejectedMutations': rejected}))
