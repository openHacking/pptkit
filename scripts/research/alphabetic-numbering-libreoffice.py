#!/usr/bin/env python3
"""Capture bounded native-PDF numbering observations; no GUI or native editing."""
import hashlib
import json
import platform
import subprocess
import tempfile
import xml.etree.ElementTree as E
import zipfile
from pathlib import Path
import fitz

ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
BASE = ROOT / '.tmp/alphabetic-numbering-libreoffice'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    BASE.mkdir(parents=True, exist_ok=True); run = Path(tempfile.mkdtemp(prefix='run-', dir=BASE)); source = run / 'numbering.pptx'
    subprocess.run(['node', 'scripts/prepare-alphabetic-numbering.mjs', str(source)], cwd=ROOT, check=True)
    with zipfile.ZipFile(source) as z:
        xml = E.fromstring(z.read('ppt/slides/slide1.xml'))
        bullets = [{'type': b.get('type'), 'startAt': int(b.get('startAt'))} for b in xml.iter(A + 'buAutoNum')]
    if bullets != [{'type': 'alphaLcPeriod', 'startAt': n} for n in [26, 27, 28, 52, 53, 54]]: raise ValueError('source numbering mismatch')
    process = subprocess.run([str(SOFFICE), f'-env:UserInstallation={(run / "profile").as_uri()}', '--headless', '--convert-to', 'pdf', '--outdir', str(run), str(source)], capture_output=True, text=True, timeout=90)
    pdf = run / 'numbering.pdf'
    if process.returncode or not pdf.exists(): raise RuntimeError(process.stderr)
    document = fitz.open(pdf)
    if len(document) != 1: raise ValueError('unexpected PDF count')
    words = document[0].get_text('words')
    pairs = [(words[i][4], words[i+1][4]) for i in range(0, len(words), 2)]
    expected = [('z.', 'VALUE26'), ('aa.', 'VALUE27'), ('ab.', 'VALUE28'), ('az.', 'VALUE52'), ('ba.', 'VALUE53'), ('bb.', 'VALUE54')]
    if pairs != expected: raise ValueError(f'Observed marker sequence changed: {pairs}')
    pixmap = document[0].get_pixmap(alpha=False)
    evidence = {'scope': 'One authored slide, six alpha-lower starts, one LibreOfficeDev headless PDF export. Extracted words and one page raster; no native editing, PowerPoint or three-round claim.',
                'application': subprocess.run([str(SOFFICE), '--version'], capture_output=True, text=True, check=True).stdout.strip(), 'platform': platform.platform(),
                'source': str(source.relative_to(ROOT)), 'sourceSha256': sha(source), 'pdf': str(pdf.relative_to(ROOT)), 'pdfSha256': sha(pdf),
                'generatorSha256': sha(ROOT / 'scripts/prepare-alphabetic-numbering.mjs'), 'runnerSha256': sha(Path(__file__)), 'sourceBullets': bullets,
                'words': [{'text': w[4], 'bbox': list(w[:4])} for w in words], 'rasterSha256': hashlib.sha256(pixmap.samples).hexdigest(), 'commandExit': process.returncode}
    (ROOT / 'docs/testing/alphabetic-numbering-libreoffice-evidence.json').write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps({'pairs': pairs, 'application': evidence['application']}, ensure_ascii=False))

if __name__ == '__main__': main()
