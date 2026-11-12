#!/usr/bin/env python3
"""Verify retained custom-property evidence against the recorded XML inventories."""
import importlib.util
import json
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
manifest = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/custom-properties-libreoffice-evidence.json'
evidence = json.loads(manifest.read_text())
def require(condition, message):
    if not condition: raise ValueError(message)
runner = ROOT / 'scripts/research/custom-properties-libreoffice.py'
spec = importlib.util.spec_from_file_location('custom_probe', runner)
probe = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe)
require(probe.digest(runner) == evidence['runnerSha256'], 'runner hash mismatch')
seed = (ROOT / evidence['seed']['path']).resolve()
require(seed.is_relative_to((ROOT / '.tmp/core-metadata-libreoffice').resolve()), 'seed outside retained directory')
require(probe.digest(seed) == evidence['seed']['sha256'], 'seed hash mismatch')
records = evidence['records']
require(len(records) == 4, 'expected four rounds')
for index, record in enumerate(records):
    require(record['round'] == index, 'round order mismatch')
    path = (ROOT / record['path']).resolve()
    require(path.is_relative_to((ROOT / '.tmp/custom-properties-libreoffice').resolve()), 'artifact outside retained directory')
    require(probe.digest(path) == record['sha256'], 'artifact hash mismatch')
    require(probe.inventory(path) == record['properties'], 'inventory mismatch')
print(json.dumps({'status': 'verified', 'records': len(records)}))
