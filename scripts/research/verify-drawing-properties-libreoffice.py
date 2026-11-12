#!/usr/bin/env python3
"""Recheck retained experiment bytes, code hashes and independent inventories."""
import hashlib
import importlib.util
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
manifest = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/drawing-properties-libreoffice-evidence.json'
evidence = json.loads(manifest.read_text())


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def safe_path(value):
    path = (ROOT / value).resolve()
    if not path.is_relative_to((ROOT / '.tmp' / 'drawing-properties-libreoffice').resolve()):
        raise ValueError(f'Evidence artifact outside retained directory: {value}')
    return path


oracle_path = ROOT / 'scripts/research/drawing-properties-oracle.py'
runner_path = ROOT / 'scripts/research/drawing-properties-libreoffice.mjs'
def require(condition, message):
    if not condition:
        raise ValueError(message)


require(evidence['schemaVersion'] == 1, 'Unsupported evidence version')
require(digest(oracle_path) == evidence['oracleSha256'], 'Oracle hash mismatch')
require(digest(runner_path) == evidence['runnerSha256'], 'Runner hash mismatch')
spec = importlib.util.spec_from_file_location('drawing_properties_oracle', oracle_path)
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)
original = safe_path(evidence['original']['path'])
require(digest(original) == evidence['original']['sha256'], 'Original file hash mismatch')
require(len(evidence['rounds']) == 4, 'Expected four rounds')
for expected_round, record in enumerate(evidence['rounds']):
    require(record['round'] == expected_round, 'Round order mismatch')
    file = safe_path(record['path'])
    require(digest(file) == record['sha256'], f'File hash mismatch: {file}')
    require(oracle.inspect(file) == record['inventory'], f'Inventory mismatch: {file}')
with zipfile.ZipFile(original) as before, zipfile.ZipFile(safe_path(evidence['rounds'][0]['path'])) as after:
    require(set(before.namelist()) == set(after.namelist()), 'Package entry sets differ')
    changed = sorted(name for name in before.namelist() if before.read(name) != after.read(name))
    require(changed == evidence['sourceChangedParts'] == ['ppt/presentation.xml', 'ppt/slides/slide1.xml'], 'Unexpected source package differences')
print(json.dumps({'status': 'verified', 'files': len(evidence['rounds']) + 1, 'rounds': 4, 'changedParts': changed}))
