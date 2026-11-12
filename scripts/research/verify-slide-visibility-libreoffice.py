#!/usr/bin/env python3
"""Verify recorded file hashes, source differences and independent inventories."""
import hashlib
import importlib.util
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
manifest = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/slide-visibility-libreoffice-evidence.json'
evidence = json.loads(manifest.read_text())


def require(condition, message):
    if not condition:
        raise ValueError(message)


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def retained(value):
    path = (ROOT / value).resolve()
    require(path.is_relative_to((ROOT / '.tmp' / 'slide-visibility-libreoffice').resolve()), f'Artifact outside retained directory: {value}')
    return path


oracle_path = ROOT / 'scripts/research/slide-visibility-oracle.py'
runner_path = ROOT / 'scripts/research/slide-visibility-libreoffice.mjs'
require(evidence['schemaVersion'] == 1, 'Unsupported evidence version')
require(digest(oracle_path) == evidence['oracleSha256'], 'Oracle hash mismatch')
require(digest(runner_path) == evidence['runnerSha256'], 'Runner hash mismatch')
spec = importlib.util.spec_from_file_location('visibility_oracle', oracle_path)
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)
records = evidence['records']
require(len(records) == 8, 'Expected four paired rounds')
for index, record in enumerate(records):
    require(record['round'] == index // 2 and record['kind'] == ('control' if index % 2 == 0 else 'hidden'), 'Round order mismatch')
    pptx, pdf = retained(record['pptx']['path']), retained(record['pdf']['path'])
    require(digest(pptx) == record['pptx']['sha256'] and digest(pdf) == record['pdf']['sha256'], f'Artifact hash mismatch in round {index}')
    require({'pptx': oracle.inspect_pptx(pptx), 'pdf': oracle.inspect_pdf(pdf)} == record['inventory'], f'Inventory mismatch in round {index}')
with zipfile.ZipFile(retained(records[0]['pptx']['path'])) as before, zipfile.ZipFile(retained(records[1]['pptx']['path'])) as after:
    require(set(before.namelist()) == set(after.namelist()), 'Source package entry sets differ')
    changed = sorted(name for name in before.namelist() if before.read(name) != after.read(name))
    require(changed == evidence['sourceChangedParts'] == ['ppt/slides/slide2.xml'], 'Source differences exceed approved slide')
print(json.dumps({'status': 'verified', 'records': len(records), 'files': len(records) * 2, 'sourceChangedParts': changed}))
