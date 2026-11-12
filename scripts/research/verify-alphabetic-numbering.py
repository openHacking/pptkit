#!/usr/bin/env python3
"""Read retained PPTX/PDF directly and reject stale or mismatched observations."""
import hashlib
import json
import sys
import xml.etree.ElementTree as E
import zipfile
from pathlib import Path
import fitz

ROOT = Path(__file__).resolve().parents[2]
BASE = (ROOT / '.tmp/alphabetic-numbering-libreoffice').resolve()
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def verify(manifest):
    evidence = json.loads(manifest.read_text())
    for key, name in [('generatorSha256', 'scripts/prepare-alphabetic-numbering.mjs'), ('runnerSha256', 'scripts/research/alphabetic-numbering-libreoffice.py')]:
        if evidence[key] != sha(ROOT / name): raise ValueError('producer changed')
    files = []
    for key in ['source', 'pdf']:
        file = (ROOT / evidence[key]).resolve()
        if not file.is_relative_to(BASE) or sha(file) != evidence[key + 'Sha256']: raise ValueError('artifact changed')
        files.append(file)
    with zipfile.ZipFile(files[0]) as z:
        if len(z.namelist()) != len(set(z.namelist())): raise ValueError('duplicate parts')
        root = E.fromstring(z.read('ppt/slides/slide1.xml'))
        bullets = [{'type': b.get('type'), 'startAt': int(b.get('startAt'))} for b in root.iter(A + 'buAutoNum')]
        if bullets != evidence['sourceBullets'] or bullets != [{'type': 'alphaLcPeriod', 'startAt': n} for n in [26, 27, 28, 52, 53, 54]]: raise ValueError('source bullets differ')
        if [n.text for n in root.iter(A + 't')] != [f'VALUE{n}' for n in [26, 27, 28, 52, 53, 54]]: raise ValueError('source label differs')
    pdf = fitz.open(files[1]);
    if len(pdf) != 1: raise ValueError('PDF pages changed')
    words = pdf[0].get_text('words'); actual = [{'text': w[4], 'bbox': list(w[:4])} for w in words]
    if actual != evidence['words'] or [w['text'] for w in actual] != [v for n, marker in [(26, 'z.'), (27, 'aa.'), (28, 'ab.'), (52, 'az.'), (53, 'ba.'), (54, 'bb.')] for v in (marker, f'VALUE{n}')]: raise ValueError('PDF labels differ')
    if hashlib.sha256(pdf[0].get_pixmap(alpha=False).samples).hexdigest() != evidence['rasterSha256'] or evidence['commandExit'] != 0: raise ValueError('PDF raster or execution differs')
    return {'verified': True, 'markers': 6, 'pages': 1}
if __name__ == '__main__': print(json.dumps(verify(Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/alphabetic-numbering-libreoffice-evidence.json')))
