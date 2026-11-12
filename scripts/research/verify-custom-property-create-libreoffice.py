#!/usr/bin/env python3
"""Independently verify the recorded custom-property creation experiment."""
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
MIME = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'

def digest(data): return hashlib.sha256(data).hexdigest()
def inspect(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or len(z.namelist()) != len(set(z.namelist())): raise ValueError('invalid ZIP')
        matches = [r for r in ET.fromstring(z.read('_rels/.rels')) if r.get('Type') == REL]
        if not matches: return []
        if len(matches) != 1 or matches[0].get('TargetMode', 'Internal') != 'Internal': raise ValueError('relationship invalid')
        part = matches[0].get('Target', '').lstrip('/')
        if not part or '..' in part.split('/'): raise ValueError('target invalid')
        root = ET.fromstring(z.read(part)); types = ET.fromstring(z.read('[Content_Types].xml'))
        if root.tag != f'{{{CUSTOM}}}Properties' or not any(e.tag == f'{{{CT}}}Override' and e.get('PartName') == '/' + part and e.get('ContentType') == MIME for e in types): raise ValueError('part declaration invalid')
        entries = []
        for child in root:
            if child.tag != f'{{{CUSTOM}}}property' or len(child) != 1: raise ValueError('unexpected property')
            entries.append({'name': child.get('name'), 'pid': child.get('pid'), 'type': child[0].tag.split('}')[-1], 'value': ''.join(child[0].itertext())})
        return entries

def main():
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/custom-property-create-libreoffice-evidence.json'
    evidence = json.loads(path.read_text()); records = evidence['records']
    if digest((ROOT / 'scripts/research/custom-property-create-libreoffice.py').read_bytes()) != evidence['runnerSha256']: raise ValueError('runner hash mismatch')
    expected_logs = {(n, kind) for n in (1, 2, 3) for kind in ('control', 'edited')}
    if len(records) != 8 or len(evidence['logs']) != 6 or {(log['round'], log['kind']) for log in evidence['logs']} != expected_logs or any(log['exitCode'] != 0 for log in evidence['logs']): raise ValueError('incomplete native rounds')
    if evidence['sourceChangedParts'] != ['[Content_Types].xml','_rels/.rels'] or evidence['sourceAddedParts'] != ['pptkit-custom-properties-1.xml']: raise ValueError('wrong source delta')
    expected = {(0,'control'): [], (0,'edited'): [('Text','2','lpwstr','中文 & text'),('Boolean','3','bool','true'),('Integer','4','i4','42'),('Real','5','r8','3.5'),('Date','6','filetime','2020-01-02T03:04:05.000Z')]}
    for n in (1,2,3):
        expected[(n,'control')] = [('PresentationFormat','2','lpwstr','On-screen Show (16:9)'),('Slides','3','r8','1')]
        expected[(n,'edited')] = [('Boolean','2','bool','1'),('Date','3','filetime','2020-01-02T03:04:05Z'),('Integer','4','r8','42'),('PresentationFormat','5','lpwstr','On-screen Show (16:9)'),('Real','6','r8','3.5'),('Slides','7','r8','1'),('Text','8','lpwstr','中文 & text')]
    observed = set(); paths = {}
    for record in records:
        key = (record['round'], record['kind']); p = (ROOT / record['path']).resolve()
        if not p.is_relative_to((ROOT / '.tmp/custom-property-create-libreoffice').resolve()) or key in observed or key not in expected or digest(p.read_bytes()) != record['sha256']: raise ValueError('record identity mismatch')
        paths[key] = p
        observed.add(key); actual = inspect(p)
        if actual != record['properties'] or [(x['name'],x['pid'],x['type'],x['value']) for x in actual] != expected[key]: raise ValueError('property inventory mismatch')
    a = paths[(0, 'control')]; b = paths[(0, 'edited')]
    with zipfile.ZipFile(a) as before, zipfile.ZipFile(b) as after:
        if sorted(set(after.namelist()) - set(before.namelist())) != evidence['sourceAddedParts'] or set(before.namelist()) - set(after.namelist()) or sorted(n for n in before.namelist() if before.read(n) != after.read(n)) != evidence['sourceChangedParts']: raise ValueError('actual source delta mismatch')
    print('Verified 8 inventories, source ZIP delta, hashes and 6 headless resaves.')
if __name__ == '__main__': main()
