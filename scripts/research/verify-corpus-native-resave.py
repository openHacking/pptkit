#!/usr/bin/env python3
"""Verify the retained artifacts and recorded comparisons of a native resave."""
import argparse
import importlib.util
import json
from pathlib import Path

import fitz

spec = importlib.util.spec_from_file_location('native_resave', Path(__file__).with_name('corpus-native-resave.py'))
resave = importlib.util.module_from_spec(spec)
spec.loader.exec_module(resave)
ROOT = resave.ROOT
render = resave.render


def verify(evidence):
    if evidence['schemaVersion'] != 1 or evidence['pymupdfVersion'] != fitz.VersionBind:
        raise ValueError('Unsupported evidence schema or different PyMuPDF version')
    if render.sha(Path(resave.__file__)) != evidence['runnerSha256'] or render.sha(Path(render.__file__)) != evidence['rasterAuditSha256']:
        raise ValueError('Native resave or raster audit implementation differs')
    pin = ROOT / 'docs/testing/libreoffice-evidence.json'
    if render.sha(pin) != evidence['pinnedCorpusEvidenceSha256']:
        raise ValueError('Corpus pin differs')
    original = {item['fixture']: item['input'] for item in json.loads(pin.read_text())['fixtures']}
    directory = (ROOT / evidence['directory']).resolve()
    if directory.parent != ROOT / '.tmp' or not directory.name.startswith('corpus-native-resave-'):
        raise ValueError('Unexpected corpus output directory')
    if [item['fixture'] for item in evidence['fixtures']] != list(original):
        raise ValueError('Evidence fixture order or coverage differs')
    for item in evidence['fixtures']:
        name = item['fixture']
        source = ROOT / '.tmp/fixtures' / f'{name}.pptx'
        render.retained.verified(source, original[name])
        for path, label in [(source, 'source'), (directory / 'saved' / f'{name}.pptx', 'saved')]:
            if resave.inspect(path) != item[label]:
                raise ValueError(f'{name}: {label} package inventory differs')
        paths = [directory / folder / f'{name}.pdf' for folder in ('source-pdf', 'reopened-pdf')]
        with fitz.open(paths[0]) as a, fitz.open(paths[1]) as b:
            for path, pdf, label in zip(paths, (a, b), ('sourcePdf', 'reopenedPdf')):
                if {'pages': len(pdf), 'sha256': render.sha(path)} != item[label]:
                    raise ValueError(f'{name}: {label} inventory differs')
            if [render.pixels(a[i], b[i]) for i in range(min(len(a), len(b)))] != item['pairedPageRasters']:
                raise ValueError(f'{name}: paired raster metrics differ')
    return len(evidence['fixtures'])


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('evidence', type=Path)
    args = parser.parse_args()
    print(f'Verified all PPTX/PDF artifacts and raster comparisons for {verify(json.loads(args.evidence.read_text()))} fixtures.')
