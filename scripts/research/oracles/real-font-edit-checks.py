"""Actual ZIP/PDF negative controls for named SC native editing evidence."""
import importlib.util, json, sys, tempfile, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec = importlib.util.spec_from_file_location('edit', Path(__file__).with_name('real-font-edit.py'))
o = importlib.util.module_from_spec(spec); spec.loader.exec_module(o)
p = Path(sys.argv[1]); expected = json.loads((p/'expected.json').read_text()); source = o.read_package(p/'fresh-round3.pptx')
rejected = []
def reject(name, fn):
    try: fn()
    except AssertionError: rejected.append(name); return
    raise AssertionError('Accepted mutation: '+name)

def rewrite(file, output, change=None, theme=None):
    with zipfile.ZipFile(file) as z, zipfile.ZipFile(output,'w') as out:
        for entry in z.infolist():
            data = z.read(entry.filename)
            if change and entry.filename == 'ppt/slides/slide2.xml':
                root = ET.fromstring(data)
                body = root.findall('p:cSld/p:spTree/p:sp',o.NS)[1].find('p:txBody',o.NS)
                change(body); data = ET.tostring(root)
            if theme is not None and entry.filename == 'ppt/theme/theme1.xml': data = theme
            out.writestr(entry,data)

def run(body, index=1): return body.findall('a:p/a:r',o.NS)[index]
def style(body, index=1): return run(body,index).find('a:rPr',o.NS)
def merged(body):
    first, second = body.findall('a:p/a:r',o.NS)
    first.find('a:t',o.NS).text += second.find('a:t',o.NS).text
    body.find('a:p',o.NS).remove(second)

with tempfile.TemporaryDirectory(prefix='pptkit-edit-negative-') as temp:
    temp = Path(temp)
    cases = [
        ('text','edit',lambda b: setattr(run(b).find('a:t',o.NS),'text','Ψ')),
        ('language','edit',lambda b: style(b).set('lang','en-US')),
        ('restored-Greek-map','edit',lambda b: style(b).find('a:latin',o.NS).set('typeface','Georgia')),
        ('merged-run','edit',merged),
        ('missing-bold','bold',lambda b: style(b).set('b','0')),
        ('missing-italic','italic',lambda b: style(b).set('i','0')),
    ]
    for name, stage, change in cases:
        file = temp/(name+'.pptx'); rewrite(p/('fresh-'+stage+'-sc.pptx'),file,change)
        reject(name,lambda: o.inspect_package(file,source,stage))
    with zipfile.ZipFile(p/'fresh-round3.pptx') as z: original_theme = z.read('ppt/theme/theme1.xml')
    file = temp/'unchanged-theme.pptx'; rewrite(p/'fresh-theme-sc.pptx',file,theme=original_theme)
    reject('unchanged-theme',lambda: o.inspect_package(file,source,'theme'))
    reject('wrong-actual-bold-face',lambda: o.inspect_pdf(p/'fresh-bold-sc.pdf',expected,'edit'))
    doc = fitz.open(p/'fresh-italic-sc.pdf'); changed = False
    for xref in doc[1].get_contents():
        stream = doc.xref_stream(xref)
        if b'67.96875' in stream:
            doc.update_stream(xref,stream.replace(b'67.96875',b'0')); changed = True
    assert changed
    file = temp/'no-shear.pdf'; doc.save(file); doc.close()
    reject('italic-without-shear',lambda: o.inspect_pdf(file,expected,'italic'))
    doc = fitz.open(p/'fresh-edit-sc.pdf'); doc[1].add_redact_annot(fitz.Rect(211,60,300,200),fill=(1,1,1)); doc[1].apply_redactions()
    file = temp/'missing-omega.pdf'; doc.save(file); doc.close()
    reject('missing-appended-glyph',lambda: o.inspect_pdf(file,expected,'edit'))
print(json.dumps({'rejectedMutations':rejected}))
