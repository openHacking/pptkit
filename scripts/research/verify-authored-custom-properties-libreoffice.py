#!/usr/bin/env python3
"""Verify retained authored-custom-property experiment records."""
import hashlib
import importlib.util
import json
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
manifest = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/authored-custom-properties-libreoffice-evidence.json'
evidence = json.loads(manifest.read_text())
def require(condition, message):
    if not condition: raise ValueError(message)
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
runner = ROOT / 'scripts/research/authored-custom-properties-libreoffice.py'
oracle_path = ROOT / 'scripts/research/custom-properties-libreoffice.py'
require(digest(runner) == evidence['runnerSha256'], 'runner hash mismatch')
require(digest(oracle_path) == evidence['oracleSha256'], 'oracle hash mismatch')
spec = importlib.util.spec_from_file_location('custom_oracle', oracle_path)
oracle = importlib.util.module_from_spec(spec); spec.loader.exec_module(oracle)
require(len(evidence['records']) == 4, 'expected four rounds')
for index, record in enumerate(evidence['records']):
    require(record['round'] == index, 'round order mismatch')
    path = (ROOT / record['path']).resolve()
    require(path.is_relative_to((ROOT / '.tmp/authored-custom-properties-libreoffice').resolve()), 'artifact outside retained directory')
    require(digest(path) == record['sha256'], 'artifact hash mismatch')
    require(oracle.inventory(path) == record['properties'], 'property inventory mismatch')
print(json.dumps({'status': 'verified', 'records': len(evidence['records'])}))
