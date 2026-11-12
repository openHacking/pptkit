#!/usr/bin/env python3
"""Independently verify paired company preservation observations."""
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT = Path(__file__).resolve().parents[2]
EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
DC = 'http://purl.org/dc/elements/1.1/'
def digest(data): return hashlib.sha256(data).hexdigest()
def inspect(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or len(z.namelist()) != len(set(z.namelist())): raise ValueError('invalid ZIP')
        app = ET.fromstring(z.read('docProps/app.xml')); core = ET.fromstring(z.read('docProps/core.xml'))
        if app.tag != f'{{{EP}}}Properties' or core.tag != f'{{{CP}}}coreProperties': raise ValueError('wrong root')
        result = {}
        for name, root, namespace, tag in [('company',app,EP,'Company'),('application',app,EP,'Application'),('slides',app,EP,'Slides'),('title',core,DC,'title'),('revision',core,CP,'revision')]:
            matches = root.findall(f'{{{namespace}}}{tag}')
            if len(matches) > 1: raise ValueError('duplicate field')
            result[name] = ''.join(matches[0].itertext()) if matches else None
        return result

def main():
    manifest = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/extended-company-libreoffice-evidence.json'
    data = json.loads(manifest.read_text()); records = data['records']; logs = data['logs']
    if digest((ROOT / 'scripts/research/extended-company-libreoffice.py').read_bytes()) != data['runnerSha256']: raise ValueError('runner mismatch')
    if len(records) != 8 or len(logs) != 6 or {(x['round'], x['kind']) for x in logs} != {(n, k) for n in (1,2,3) for k in ('control','edited')} or any(x['exitCode'] != 0 for x in logs): raise ValueError('incomplete rounds')
    if data['sourceChangedParts'] != ['docProps/app.xml']: raise ValueError('source change declaration mismatch')
    paths = {}
    for record in records:
        key = (record['round'], record['kind']); path = (ROOT / record['path']).resolve()
        if key in paths or key not in {(n, k) for n in (0,1,2,3) for k in ('control','edited')} or not path.is_relative_to((ROOT / '.tmp/extended-company-libreoffice').resolve()) or digest(path.read_bytes()) != record['sha256']: raise ValueError('record identity mismatch')
        actual = inspect(path); expected = {'company':'Old Company' if key[1] == 'control' else 'New & 中文','application':'PPTKit' if key[0] == 0 else 'LibreOfficeDev/26.8.0.0.alpha0$MacOSX_AARCH64 LibreOffice_project/2c87e51eeaa2b413ff4ae097b2705eea1995d8e5','slides':'1' if key[0] == 0 else None,'title':'Company control','revision':'7'}
        if actual != expected or record['metadata'] != actual: raise ValueError('metadata mismatch')
        paths[key] = path
    with zipfile.ZipFile(paths[(0,'control')]) as before, zipfile.ZipFile(paths[(0,'edited')]) as after:
        if set(before.namelist()) != set(after.namelist()) or sorted(n for n in before.namelist() if before.read(n) != after.read(n)) != data['sourceChangedParts']: raise ValueError('source ZIP mismatch')
        old, new = before.read('docProps/app.xml'), after.read('docProps/app.xml'); marker = b'<Company>Old Company</Company>'; replacement = '<Company>New &amp; 中文</Company>'.encode('utf-8')
        if old.count(marker) != 1 or old.replace(marker, replacement) != new: raise ValueError('source company edit was not isolated')
    print('Verified 8 metadata inventories, isolated company edit, hashes and 6 headless resaves.')
if __name__ == '__main__': main()
