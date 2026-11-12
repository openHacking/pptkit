#!/usr/bin/env python3
"""Resave the pinned REL-02 corpus with a selected native Impress executable."""
import argparse
import importlib.util
import json
import platform
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import fitz

spec = importlib.util.spec_from_file_location('render_audit', Path(__file__).with_name('corpus-headless-render-audit.py'))
render = importlib.util.module_from_spec(spec)
spec.loader.exec_module(render)
ROOT = render.ROOT
CHART = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'


def inspect(path):
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad is not None:
            raise ValueError(f'ZIP CRC failure: {path}: {bad}')
        names = archive.namelist()
        if len(set(names)) != len(names):
            raise ValueError(f'Duplicate ZIP entry: {path}')
        xml = {name: ET.fromstring(archive.read(name)) for name in names if name.endswith(('.xml', '.rels'))}
        charts, external_data = {}, {}
        for name, tree in sorted(xml.items()):
            if not re.fullmatch(r'ppt/charts/chart\d+\.xml', name):
                continue
            external_data[name] = [node.attrib for node in tree.iter(f'{CHART}externalData')]
            charts[name] = [
                [[point.attrib['idx'], point.findtext(f'{CHART}v')] for point in cache.findall(f'{CHART}pt')]
                for cache in tree.iter(f'{CHART}numCache')
            ]
        return {'sha256': render.sha(path), 'bytes': path.stat().st_size, 'xmlParts': len(xml),
                'slides': render.retained.slides(path), 'chartNumericCaches': charts, 'chartExternalDataAttributes': external_data,
                'embeddedXlsxParts': sorted(name for name in names if name.startswith('ppt/embeddings/') and name.endswith('.xlsx'))}


def run(executable, output):
    pin = ROOT / 'docs/testing/libreoffice-evidence.json'
    records = json.loads(pin.read_text())['fixtures']
    sources = [ROOT / '.tmp/fixtures' / f"{item['fixture']}.pptx" for item in records]
    for item, source in zip(records, sources):
        render.retained.verified(source, item['input'])
    version = subprocess.run([executable, '--version'], capture_output=True, text=True, timeout=30, check=True).stdout.strip()
    if not version:
        raise ValueError('Application version was empty')
    work = Path(tempfile.mkdtemp(prefix='corpus-native-resave-', dir=ROOT / '.tmp')).resolve()
    saved, before, after = [work / name for name in ('saved', 'source-pdf', 'reopened-pdf')]
    for folder in (saved, before, after):
        folder.mkdir()
    save = subprocess.run([executable, f'-env:UserInstallation={(work / "profile-save").as_uri()}',
                           '--headless', '--convert-to', 'pptx:Impress MS PowerPoint 2007 XML',
                           '--outdir', str(saved), *map(str, sources)],
                          capture_output=True, text=True, timeout=120, check=True)
    targets = [saved / source.name for source in sources]
    inputs, outputs = [inspect(path) for path in sources], [inspect(path) for path in targets]
    logs = {'save': {'stdout': save.stdout, 'stderr': save.stderr},
            'sourcePdf': render.convert(executable, work / 'profile-source', before, sources),
            'reopenedPdf': render.convert(executable, work / 'profile-reopen', after, targets)}
    fixtures = []
    for source, initial, final in zip(sources, inputs, outputs):
        source_pdf, reopened_pdf = [folder / f'{source.stem}.pdf' for folder in (before, after)]
        with fitz.open(source_pdf) as a, fitz.open(reopened_pdf) as b:
            fixtures.append({'fixture': source.stem, 'source': initial, 'saved': final,
                             'sourcePdf': {'pages': len(a), 'sha256': render.sha(source_pdf)},
                             'reopenedPdf': {'pages': len(b), 'sha256': render.sha(reopened_pdf)},
                             'pairedPageRasters': [render.pixels(a[i], b[i]) for i in range(min(len(a), len(b)))]})
    result = {'schemaVersion': 1, 'scope': 'Native headless save/reopen and paired PDF raster observation; no interactive editing, repair-dialog or calibrated visual parity claim.',
              'application': version, 'platform': platform.platform(), 'pymupdfVersion': fitz.VersionBind,
              'pinnedCorpusEvidenceSha256': render.sha(pin), 'runnerSha256': render.sha(Path(__file__)),
              'rasterAuditSha256': render.sha(Path(render.__file__)), 'directory': str(work.relative_to(ROOT)),
              'logs': logs, 'fixtures': fixtures}
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    for item in fixtures:
        print(item['fixture'], item['sourcePdf']['pages'], '->', item['reopenedPdf']['pages'],
              'pages;', len(item['source']['embeddedXlsxParts']), '->', len(item['saved']['embeddedXlsxParts']), 'embedded XLSX')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('executable', help='path to the native soffice executable')
    parser.add_argument('output', type=Path, help='path to the JSON observation')
    args = parser.parse_args()
    run(args.executable, args.output)
