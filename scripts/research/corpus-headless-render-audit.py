#!/usr/bin/env python3
"""Compare fresh LibreOffice headless renders of hash-verified corpus artifacts."""
import hashlib
import importlib.util
import json
import platform
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('retained_audit', Path(__file__).with_name('corpus-retained-text-audit.py'))
retained = importlib.util.module_from_spec(spec)
spec.loader.exec_module(retained)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def convert(executable, profile, output, inputs):
    command = [executable, f'-env:UserInstallation={profile.as_uri()}', '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', str(output), *map(str, inputs)]
    result = subprocess.run(command, capture_output=True, text=True, timeout=120, check=True)
    return {'stdout': result.stdout, 'stderr': result.stderr}


def pixels(left, right):
    if tuple(left.rect) != tuple(right.rect):
        return {'status': 'different-page-rectangles', 'source': list(left.rect), 'saved': list(right.rect)}
    a, b = left.get_pixmap(alpha=False), right.get_pixmap(alpha=False)
    if (a.width, a.height, a.n) != (b.width, b.height, b.n):
        return {'status': 'different-raster-dimensions', 'source': [a.width, a.height, a.n], 'saved': [b.width, b.height, b.n]}
    width, channels = a.width, a.n
    aa, bb = a.samples, b.samples
    count = 0
    max_delta = 0
    bounds = [a.width, a.height, -1, -1]
    for index in range(0, len(aa), channels):
        delta = max(abs(aa[index + c] - bb[index + c]) for c in range(channels))
        if delta:
            count += 1
            max_delta = max(max_delta, delta)
            x, y = (index // channels) % width, (index // channels) // width
            bounds = [min(bounds[0], x), min(bounds[1], y), max(bounds[2], x), max(bounds[3], y)]
    return {'status': 'compared', 'width': a.width, 'height': a.height, 'changedPixels': count, 'totalPixels': a.width * a.height, 'maxChannelDelta': max_delta, 'changedBoundsInclusive': bounds if count else None}


def main():
    executable = sys.argv[1] if len(sys.argv) > 1 else 'soffice'
    evidence = ROOT / 'docs/testing/libreoffice-evidence.json'
    records = json.loads(evidence.read_text())['fixtures']
    sources = [ROOT / '.tmp/fixtures' / f"{item['fixture']}.pptx" for item in records]
    saved = [ROOT / '.tmp/compatibility/roundtrip' / f"{item['fixture']}.pptx" for item in records]
    for item, source, target in zip(records, sources, saved):
        retained.verified(source, item['input'])
        retained.verified(target, item['saved'])
    version = subprocess.run([executable, '--version'], capture_output=True, text=True, timeout=30, check=True).stdout.strip()
    work = Path(tempfile.mkdtemp(prefix='corpus-headless-render-', dir=ROOT / '.tmp')).resolve()
    source_dir, saved_dir = work / 'source', work / 'saved'
    source_dir.mkdir(); saved_dir.mkdir()
    logs = {'source': convert(executable, work / 'profile-source', source_dir, sources), 'saved': convert(executable, work / 'profile-saved', saved_dir, saved)}
    fixtures = []
    for item, source, target in zip(records, sources, saved):
        name = item['fixture']; source_pdf, saved_pdf = source_dir / f'{name}.pdf', saved_dir / f'{name}.pdf'
        with fitz.open(source_pdf) as a, fitz.open(saved_pdf) as b:
            pages = [pixels(a[i], b[i]) for i in range(min(len(a), len(b)))]
            fixtures.append({'fixture': name, 'sourcePptxSha256': sha(source), 'savedPptxSha256': sha(target), 'sourcePdfSha256': sha(source_pdf), 'savedPdfSha256': sha(saved_pdf), 'historicalSavedPdfSha256': item['reopenedPdf']['sha256'], 'sourcePages': len(a), 'savedPages': len(b), 'pages': pages})
    result = {'schemaVersion': 1, 'scope': 'Fresh headless PDF render of retained source and saved PPTX; exact raster differences are observations, not interactive, semantic, repair-dialog or application support claims.', 'application': version, 'platform': platform.platform(), 'pymupdfVersion': fitz.VersionBind, 'historicalManifestSha256': sha(evidence), 'auditScriptSha256': sha(Path(__file__)), 'retainedVerifierSha256': sha(Path(retained.__file__)), 'directory': str(work.relative_to(ROOT)), 'logs': logs, 'fixtures': fixtures}
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / '.tmp/corpus-headless-render-audit.json'
    output.write_text(json.dumps(result, indent=2) + '\n')
    for item in fixtures:
        print(item['fixture'], item['sourcePages'], item['savedPages'], [(page.get('changedPixels'), page.get('totalPixels')) for page in item['pages']])


if __name__ == '__main__':
    main()
