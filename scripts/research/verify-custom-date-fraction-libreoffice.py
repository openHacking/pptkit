#!/usr/bin/env python3
"""Independently verify paired fractional-date headless observations."""
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
CUSTOM='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
NAMES=['DateTenth','DateHundredth','FileZeros','FileZero']
TYPES=['date','date','filetime','filetime']
BEFORE=['2020-01-01T00:30:00.1+08:00','2020-02-29T23:30:00.12-05:30','2021-01-01T00:30:00.1230000000+08:00','2021-02-28T23:30:00.0000-05:30']
AFTER=['2024-02-29T23:59:59.000Z','2024-02-29T23:59:59.456Z','2023-12-31T00:00:00.000Z','2023-12-31T00:00:00.789Z']
def digest(data): return hashlib.sha256(data).hexdigest()
def inspect(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or len(z.namelist())!=len(set(z.namelist())): raise ValueError('invalid ZIP')
        matches=[r for r in ET.fromstring(z.read('_rels/.rels')) if r.get('Type')==REL]
        if len(matches)!=1 or matches[0].get('TargetMode','Internal')!='Internal': raise ValueError('relationship invalid')
        name=matches[0].get('Target','').lstrip('/')
        if not name or '..' in name.split('/'): raise ValueError('target invalid')
        root=ET.fromstring(z.read(name)); result=[]
        if root.tag!=f'{{{CUSTOM}}}Properties': raise ValueError('wrong root')
        for child in root:
            if child.tag!=f'{{{CUSTOM}}}property' or len(child)!=1: raise ValueError('unexpected property')
            result.append({'name':child.get('name'),'pid':child.get('pid'),'type':child[0].tag.split('}')[-1],'value':''.join(child[0].itertext())})
        return result

def main():
    manifest=Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'docs/testing/custom-date-fraction-libreoffice-evidence.json'; data=json.loads(manifest.read_text()); records=data['records']; logs=data['logs']
    if digest((ROOT/'scripts/research/custom-date-fraction-libreoffice.py').read_bytes())!=data['runnerSha256'] or digest((ROOT/'scripts/research/custom-property-create-libreoffice.py').read_bytes())!=data['oracleSha256']: raise ValueError('runner mismatch')
    prior=json.loads((ROOT/'docs/testing/custom-property-create-libreoffice-evidence.json').read_text()); seed=next(x for x in prior['records'] if x['round']==0 and x['kind']=='edited'); seed_path=(ROOT/seed['path']).resolve()
    if not seed_path.is_relative_to((ROOT/'.tmp/custom-property-create-libreoffice').resolve()) or seed['sha256']!=data['seedSha256'] or digest(seed_path.read_bytes())!=data['seedSha256']: raise ValueError('seed mismatch')
    if len(records)!=8 or len(logs)!=6 or {(x['round'],x['kind']) for x in logs}!={(n,k) for n in (1,2,3) for k in ('control','edited')} or any(x['exitCode']!=0 for x in logs): raise ValueError('incomplete rounds')
    if data['sourceChangedParts']!=['pptkit-custom-properties-1.xml']: raise ValueError('wrong source delta')
    paths={}
    for record in records:
        key=(record['round'],record['kind']); path=(ROOT/record['path']).resolve()
        if key in paths or key not in {(n,k) for n in (0,1,2,3) for k in ('control','edited')} or not path.is_relative_to((ROOT/'.tmp/custom-date-fraction-libreoffice').resolve()) or digest(path.read_bytes())!=record['sha256']: raise ValueError('record mismatch')
        if key[0]==0:
            values=BEFORE if key[1]=='control' else AFTER
            expected=[{'name':name,'pid':str(i+2),'type':TYPES[i],'value':values[i]} for i,name in enumerate(NAMES)]
        else:
            values=['2020-03-01T05:00:00Z','2019-12-31T16:30:00Z','2021-03-01T05:00:00Z','2020-12-31T16:30:00Z'] if key[1]=='control' else ['2024-02-29T23:59:59Z','2024-02-29T23:59:59Z','2023-12-31T00:00:00Z','2023-12-31T00:00:00Z']
            ordered=['DateHundredth','DateTenth','FileZero','FileZeros']
            expected=[{'name':name,'pid':str(i+2),'type':'filetime','value':values[i]} for i,name in enumerate(ordered)]+[{'name':'PresentationFormat','pid':'6','type':'lpwstr','value':'On-screen Show (16:9)'},{'name':'Slides','pid':'7','type':'r8','value':'1'}]
        actual=inspect(path)
        if actual!=expected or actual!=record['properties']: raise ValueError('inventory mismatch')
        paths[key]=path
    with zipfile.ZipFile(paths[(0,'control')]) as before,zipfile.ZipFile(paths[(0,'edited')]) as after:
        if set(before.namelist())!=set(after.namelist()) or sorted(n for n in before.namelist() if before.read(n)!=after.read(n))!=data['sourceChangedParts']: raise ValueError('source ZIP mismatch')
        old=before.read(data['sourceChangedParts'][0]); new=after.read(data['sourceChangedParts'][0]); expected=old
        for kind,left,right in zip(TYPES,BEFORE,AFTER):
            marker=f'<vt:{kind}>{left}</vt:{kind}>'.encode(); replacement=f'<vt:{kind}>{right}</vt:{kind}>'.encode()
            if expected.count(marker)!=1: raise ValueError('source marker mismatch')
            expected=expected.replace(marker,replacement,1)
        if expected!=new: raise ValueError('unexpected source edit')
    print('Verified 8 fractional-date inventories, isolated source edits, hashes and 6 headless resaves.')
if __name__=='__main__': main()
