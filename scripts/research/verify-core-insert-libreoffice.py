#!/usr/bin/env python3
"""Independently verify retained core insertion experiment artifacts."""
import hashlib, json, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]
evidence=json.loads((ROOT/'docs/testing/core-insert-libreoffice-evidence.json').read_text())
sha=lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
assert sha(ROOT/'scripts/research/core-insert-libreoffice.py') == evidence['runnerSha256']
assert evidence['sourceChangedParts']==['docProps/core.xml']
assert len(evidence['records'])==8 and len(evidence['logs'])==6
assert all(item['exitCode']==0 for item in evidence['logs'])
expected_fields={'title':'http://purl.org/dc/elements/1.1/','creator':'http://purl.org/dc/elements/1.1/','subject':'http://purl.org/dc/elements/1.1/','description':'http://purl.org/dc/elements/1.1/','language':'http://purl.org/dc/elements/1.1/','revision':'http://schemas.openxmlformats.org/package/2006/metadata/core-properties','keywords':'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'}
seen=set()
for record in evidence['records']:
    key=(record['round'],record['kind']); assert key not in seen; seen.add(key)
    path=ROOT/record['path']; assert sha(path)==record['sha256']
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None and len(z.namelist())==len(set(z.namelist()))
        root=ET.fromstring(z.read('docProps/core.xml')); actual={}
        for field,ns in expected_fields.items():
            values=root.findall(f'{{{ns}}}{field}'); assert len(values)<=1
            actual[field]=''.join(values[0].itertext()) if values else None
        assert actual==record['metadata']
    assert actual['creator']=='Sentinel Author' and actual['revision']=='7'
    assert actual['keywords']==('alpha, beta' if key[0]==0 else 'alpha beta')
    if key[1]=='edited': assert [actual[f] for f in ('title','subject','description','language')]==['PPTKit Presentation','','Inserted & 中文','fr-FR']
    elif key[0]==0: assert all(actual[f] is None for f in ('title','subject','description','language'))
    else: assert [actual[f] for f in ('title','subject','description','language')]==['','','','zh-CN']
assert seen=={(n,k) for n in range(4) for k in ('control','edited')}
a=ROOT/next(r['path'] for r in evidence['records'] if r['round']==0 and r['kind']=='control')
b=ROOT/next(r['path'] for r in evidence['records'] if r['round']==0 and r['kind']=='edited')
with zipfile.ZipFile(a) as x, zipfile.ZipFile(b) as y:
    assert set(x.namelist())==set(y.namelist())
    assert sorted(n for n in x.namelist() if x.read(n)!=y.read(n))==['docProps/core.xml']
print('Verified eight retained artifacts, source isolation and recorded inventories.')
