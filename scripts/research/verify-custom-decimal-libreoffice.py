#!/usr/bin/env python3
"""Independently verify paired decimal headless observations and source edits."""
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT = Path(__file__).resolve().parents[2]
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
NAMES = ['Normal','Fraction','Inexact','Large']
BEFORE = ['1.25','0.125','0.1','9007199254740992']
AFTER = ['2.5','0.5','0.1','4503599627370496']
def digest(data): return hashlib.sha256(data).hexdigest()
def inspect(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or len(z.namelist()) != len(set(z.namelist())): raise ValueError('invalid ZIP')
        matches=[r for r in ET.fromstring(z.read('_rels/.rels')) if r.get('Type')==REL]
        if len(matches)!=1 or matches[0].get('TargetMode','Internal')!='Internal': raise ValueError('relationship invalid')
        part=matches[0].get('Target','').lstrip('/')
        if not part or '..' in part.split('/'): raise ValueError('target invalid')
        root=ET.fromstring(z.read(part))
        if root.tag!=f'{{{CUSTOM}}}Properties': raise ValueError('wrong root')
        result=[]
        for child in root:
            if child.tag!=f'{{{CUSTOM}}}property' or len(child)!=1: raise ValueError('unexpected property')
            result.append({'name':child.get('name'),'pid':child.get('pid'),'type':child[0].tag.split('}')[-1],'value':''.join(child[0].itertext())})
        return result

def main():
    path=Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'docs/testing/custom-decimal-libreoffice-evidence.json'; data=json.loads(path.read_text()); records=data['records']; logs=data['logs']
    if digest((ROOT/'scripts/research/custom-decimal-libreoffice.py').read_bytes())!=data['runnerSha256'] or digest((ROOT/'scripts/research/custom-property-create-libreoffice.py').read_bytes())!=data['oracleSha256']: raise ValueError('runner mismatch')
    prior=json.loads((ROOT/'docs/testing/custom-property-create-libreoffice-evidence.json').read_text()); seed=next(x for x in prior['records'] if x['round']==0 and x['kind']=='edited'); seed_path=(ROOT/seed['path']).resolve()
    if not seed_path.is_relative_to((ROOT/'.tmp/custom-property-create-libreoffice').resolve()) or seed['sha256']!=data['seedSha256'] or digest(seed_path.read_bytes())!=data['seedSha256']: raise ValueError('seed mismatch')
    if len(records)!=8 or len(logs)!=6 or {(x['round'],x['kind']) for x in logs}!={(n,k) for n in (1,2,3) for k in ('control','edited')} or any(x['exitCode']!=0 for x in logs): raise ValueError('incomplete rounds')
    if data['sourceChangedParts']!=['pptkit-custom-properties-1.xml']: raise ValueError('wrong delta')
    paths={}
    for record in records:
        key=(record['round'],record['kind']); file=(ROOT/record['path']).resolve()
        if key in paths or key not in {(n,k) for n in (0,1,2,3) for k in ('control','edited')} or not file.is_relative_to((ROOT/'.tmp/custom-decimal-libreoffice').resolve()) or digest(file.read_bytes())!=record['sha256']: raise ValueError('record mismatch')
        actual=inspect(file)
        if key[0]==0:
            values=BEFORE if key[1]=='control' else AFTER
            expected=[{'name':name,'pid':str(i+2),'type':'decimal','value':values[i]} for i,name in enumerate(NAMES)]
        else:
            expected=[{'name':'PresentationFormat','pid':'2','type':'lpwstr','value':'On-screen Show (16:9)'},{'name':'Slides','pid':'3','type':'r8','value':'1'}]
        if actual!=expected or actual!=record['properties']: raise ValueError('inventory mismatch')
        paths[key]=file
    with zipfile.ZipFile(paths[(0,'control')]) as before,zipfile.ZipFile(paths[(0,'edited')]) as after:
        if set(before.namelist())!=set(after.namelist()) or sorted(n for n in before.namelist() if before.read(n)!=after.read(n))!=data['sourceChangedParts']: raise ValueError('source ZIP mismatch')
        old=before.read(data['sourceChangedParts'][0]); new=after.read(data['sourceChangedParts'][0]); expected=old
        for left,right in zip(BEFORE,AFTER):
            marker=f'<vt:decimal>{left}</vt:decimal>'.encode(); replacement=f'<vt:decimal>{right}</vt:decimal>'.encode()
            if expected.count(marker)!=1: raise ValueError('source marker mismatch')
            expected=expected.replace(marker,replacement,1)
        if expected!=new: raise ValueError('unexpected source edit')
    print('Verified 8 decimal inventories, exact source edits, hashes and 6 headless resaves.')
if __name__=='__main__': main()
