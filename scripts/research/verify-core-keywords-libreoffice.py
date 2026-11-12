#!/usr/bin/env python3
"""Independently verify the recorded paired keyword observation."""
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT = Path(__file__).resolve().parents[2]
CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
DC = 'http://purl.org/dc/elements/1.1/'
def digest(data): return hashlib.sha256(data).hexdigest()
def inspect(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or len(z.namelist()) != len(set(z.namelist())): raise ValueError('invalid ZIP')
        root = ET.fromstring(z.read('docProps/core.xml'))
        if root.tag != f'{{{CP}}}coreProperties': raise ValueError('wrong root')
        result = {}
        for field, namespace in [('title', DC), ('creator', DC), ('revision', CP), ('keywords', CP)]:
            matches = root.findall(f'{{{namespace}}}{field}')
            if len(matches) != 1: raise ValueError('missing or duplicate field')
            result[field] = ''.join(matches[0].itertext())
        return result

def main():
    manifest = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/core-keywords-libreoffice-evidence.json'
    data = json.loads(manifest.read_text()); records = data['records']; logs = data['logs']
    if digest((ROOT / 'scripts/research/core-keywords-libreoffice.py').read_bytes()) != data['runnerSha256']: raise ValueError('runner mismatch')
    if len(records) != 8 or len(logs) != 6 or {(x['round'], x['kind']) for x in logs} != {(n, k) for n in (1,2,3) for k in ('control','edited')} or any(x['exitCode'] != 0 for x in logs): raise ValueError('incomplete rounds')
    if data['sourceChangedParts'] != ['docProps/core.xml']: raise ValueError('source change declaration mismatch')
    paths = {}
    for record in records:
        key = (record['round'], record['kind']); path = (ROOT / record['path']).resolve()
        if key in paths or key not in {(n, k) for n in (0,1,2,3) for k in ('control','edited')} or not path.is_relative_to((ROOT / '.tmp/core-keywords-libreoffice').resolve()) or digest(path.read_bytes()) != record['sha256']: raise ValueError('record identity mismatch')
        actual = inspect(path); expected = {'title':'Original title','creator':'Sentinel Author','revision':'7','keywords': ('alpha, beta' if key[0] == 0 else 'alpha beta') if key[1] == 'control' else ('New, value & 中文' if key[0] == 0 else 'New value & 中文')}
        if actual != expected or record['metadata'] != actual: raise ValueError('metadata mismatch')
        paths[key] = path
    with zipfile.ZipFile(paths[(0,'control')]) as before, zipfile.ZipFile(paths[(0,'edited')]) as after:
        if set(before.namelist()) != set(after.namelist()) or sorted(n for n in before.namelist() if before.read(n) != after.read(n)) != data['sourceChangedParts']: raise ValueError('source ZIP mismatch')
        old, new = before.read('docProps/core.xml'), after.read('docProps/core.xml')
        marker = b'<cp:keywords>alpha, beta</cp:keywords>'
        replacement = '<cp:keywords>New, value &amp; 中文</cp:keywords>'.encode('utf-8')
        if old.count(marker) != 1 or old.replace(marker, replacement) != new: raise ValueError('source keyword edit was not isolated')
    print('Verified 8 metadata inventories, isolated keyword edit, hashes and 6 headless resaves.')
if __name__ == '__main__': main()
