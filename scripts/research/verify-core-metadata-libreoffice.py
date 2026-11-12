#!/usr/bin/env python3
"""Verify retained metadata experiment artifacts with an independent XML inventory."""
import hashlib
import importlib.util
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
manifest = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/core-metadata-libreoffice-evidence.json'
evidence = json.loads(manifest.read_text())

def require(condition, message):
    if not condition:
        raise ValueError(message)

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def retained(value):
    path = (ROOT / value).resolve()
    require(path.is_relative_to((ROOT / '.tmp/core-metadata-libreoffice').resolve()), f'Artifact outside retained directory: {value}')
    return path

oracle_path = ROOT / 'scripts/research/core-metadata-oracle.py'
runner_path = ROOT / 'scripts/research/core-metadata-libreoffice.mjs'
require(evidence['schemaVersion'] == 1, 'Unsupported evidence version')
require(digest(oracle_path) == evidence['oracleSha256'], 'Oracle hash mismatch')
require(digest(runner_path) == evidence['runnerSha256'], 'Runner hash mismatch')
spec = importlib.util.spec_from_file_location('metadata_oracle', oracle_path)
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)
generated = retained(evidence['generated']['path'])
require(digest(generated) == evidence['generated']['sha256'], 'Generated artifact hash mismatch')
records = evidence['records']
require(len(records) == 8, 'Expected four paired rounds')
for index, record in enumerate(records):
    require(record['round'] == index // 2 and record['kind'] == ('control' if index % 2 == 0 else 'edited'), 'Round order mismatch')
    pptx = retained(record['path'])
    require(digest(pptx) == record['sha256'], f'Artifact hash mismatch in record {index}')
    require(oracle.inventory(pptx) == record['inventory'], f'Inventory mismatch in record {index}')
with zipfile.ZipFile(retained(records[0]['path'])) as before, zipfile.ZipFile(retained(records[1]['path'])) as after:
    require(set(before.namelist()) == set(after.namelist()), 'Source package entry sets differ')
    changed = sorted(name for name in before.namelist() if before.read(name) != after.read(name))
    require(changed == evidence['sourceChangedParts'] == ['docProps/core.xml'], 'Source differences exceed core properties')
print(json.dumps({'status': 'verified', 'records': len(records), 'files': len(records) + 1, 'sourceChangedParts': changed}))
