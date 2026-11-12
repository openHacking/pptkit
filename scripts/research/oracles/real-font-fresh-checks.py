"""Reject actual ZIP/PDF corruptions; never mutate the recorded native artifacts."""
import importlib.util, json, tempfile, zipfile
from pathlib import Path
import fitz
spec = importlib.util.spec_from_file_location('oracle', Path(__file__).with_name('real-font-fresh.py'))
o = importlib.util.module_from_spec(spec); spec.loader.exec_module(o)
directory = Path(__import__('sys').argv[1]); expected = json.loads((directory/'expected.json').read_text())
def reject(fn):
    try: fn()
    except AssertionError: return
    raise AssertionError('mutated artifact was accepted')
with tempfile.TemporaryDirectory(prefix='pptkit-fresh-negative-') as temp:
    temp = Path(temp)
    for name, transform in [
        ('body', lambda s: s.replace('骨直令返', '骨直令錯')),
        ('slot', lambda s: s.replace('typeface="Hiragino Sans GB"', 'typeface="Arial"')),
        ('label', lambda s: s.replace('direct-sc', 'map-sc')),
    ]:
        file = temp/(name+'.pptx')
        with zipfile.ZipFile(directory/'fresh-fonts.pptx') as source, zipfile.ZipFile(file, 'w') as out:
            for entry in source.infolist():
                data = source.read(entry.filename)
                if entry.filename == 'ppt/slides/slide1.xml': data = transform(data.decode()).encode()
                out.writestr(entry, data)
        reject(lambda: o.inspect_package(file, expected))
    # A real reordering is rejected from PDF labels, independently of manifest order.
    doc = fitz.open(directory/'fresh-fonts.pdf'); doc.select([1,0]+list(range(2,16)))
    file = temp/'reordered.pdf'; doc.save(file); doc.close()
    reject(lambda: o.inspect_pdf(file, expected))
    doc = fitz.open(directory/'fresh-fonts.pdf')
    doc[0].add_redact_annot(fitz.Rect(0,60,720,260), fill=(1,1,1)); doc[0].apply_redactions()
    file = temp/'blank-body.pdf'; doc.save(file); doc.close()
    reject(lambda: o.inspect_pdf(file, expected))
print(json.dumps({'rejectedMutations': ['ZIP body', 'ZIP font slot', 'ZIP label', 'PDF order', 'PDF blank body']}))
