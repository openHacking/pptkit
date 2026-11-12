#!/usr/bin/env python3
"""Verify retained paired custom-property save experiment."""
import hashlib
import importlib.util
import json
import sys
import zipfile
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
manifest = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/custom-property-rename-libreoffice-evidence.json'
evidence = json.loads(manifest.read_text())
def require(condition, message):
    if not condition: raise ValueError(message)
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def retained(value):
    path = (ROOT / value).resolve()
    require(path.is_relative_to((ROOT / '.tmp/custom-property-rename-libreoffice').resolve()), 'artifact outside retained directory')
    return path
runner = ROOT / 'scripts/research/custom-property-rename-libreoffice.py'
oracle_path = ROOT / 'scripts/research/custom-properties-libreoffice.py'
require(digest(runner) == evidence['runnerSha256'], 'runner hash mismatch')
require(digest(oracle_path) == evidence['oracleSha256'], 'oracle hash mismatch')
spec = importlib.util.spec_from_file_location('custom_oracle', oracle_path)
oracle = importlib.util.module_from_spec(spec); spec.loader.exec_module(oracle)
records = evidence['records']; require(len(records) == 8, 'expected four paired rounds')
for i, item in enumerate(records):
    require(item['round'] == i // 2 and item['kind'] == ('control' if i % 2 == 0 else 'edited'), 'round order mismatch')
    path = retained(item['path']); require(digest(path) == item['sha256'], 'artifact hash mismatch')
    require(oracle.inventory(path) == item['properties'], 'inventory mismatch')
with zipfile.ZipFile(retained(records[0]['path'])) as before, zipfile.ZipFile(retained(records[1]['path'])) as after:
    require(set(before.namelist()) == set(after.namelist()), 'source entry sets differ')
    changed = sorted(name for name in before.namelist() if before.read(name) != after.read(name))
    require(changed == evidence['sourceChangedParts'] == ['docProps/custom.xml'], 'source difference mismatch')
print(json.dumps({'status':'verified','records':len(records),'changed':changed}))
