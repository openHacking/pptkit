#!/usr/bin/env python3
"""Independently verify public-authoring property samples and application outputs."""
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]
BASE=(ROOT/'.tmp/fresh-properties-libreoffice').resolve()
CUSTOM='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
GUID='{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'
def digest(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def inspect(file):
    with zipfile.ZipFile(file) as z:
        if z.testzip() is not None or len(z.namelist())!=len(set(z.namelist())):raise ValueError('invalid zip')
        relations=[e for e in ET.fromstring(z.read('_rels/.rels')) if e.get('Type')==REL]
        overrides=[e for e in ET.fromstring(z.read('[Content_Types].xml')) if e.tag==f'{{{CT}}}Override' and e.get('PartName')=='/docProps/custom.xml']
        present='docProps/custom.xml' in z.namelist()
        if len(relations)!=int(present) or len(overrides)!=int(present):raise ValueError('custom relationship or override mismatch')
        if present and (relations[0].get('Target','').lstrip('/')!='docProps/custom.xml' or relations[0].get('TargetMode','Internal')!='Internal' or overrides[0].get('ContentType')!='application/vnd.openxmlformats-officedocument.custom-properties+xml'):raise ValueError('invalid custom target')
        values=[]
        if present:
            root=ET.fromstring(z.read('docProps/custom.xml'))
            if root.tag!=f'{{{CUSTOM}}}Properties':raise ValueError('invalid root')
            for item in root:
                if item.tag!=f'{{{CUSTOM}}}property' or len(item)!=1:raise ValueError('invalid property')
                v=item[0];values.append({'name':item.get('name'),'pid':item.get('pid'),'fmtid':item.get('fmtid'),'type':v.tag.split('}')[-1],'value':v.text or ''})
        if len({x['name'] for x in values})!=len(values) or len({(x['fmtid'],x['pid']) for x in values})!=len(values):raise ValueError('duplicate property')
        return {'part_present':present,'custom_relationships':len(relations),'properties':values}
def main():
    file=Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'docs/testing/fresh-properties-libreoffice-evidence.json';data=json.loads(file.read_text());records=data['records'];logs=data['logs'];aux=data['auxiliaries'];expected={(kind,n) for kind in ('control','edited') for n in range(4)}
    if len(records)!=8 or len(logs)!=6 or {(x['kind'],x['round']) for x in logs}!={(kind,n) for kind in ('control','edited') for n in (1,2,3)} or any(x['exit_code']!=0 for x in logs):raise ValueError('incomplete rounds')
    paths={}
    for record in records+aux:
        path=(ROOT/record['path']).resolve()
        if not path.is_relative_to(BASE) or digest(path)!=record['sha256']:raise ValueError('hash or path mismatch')
        actual=inspect(path)
        if any(actual[k]!=record[k] for k in ('part_present','custom_relationships','properties')):raise ValueError('inventory mismatch')
        key=(record['kind'],record.get('round'))
        if key in paths:raise ValueError('duplicate record')
        paths[key]=path
        names=[x['name'] for x in actual['properties']]
        if key==('empty',None) and (not actual['part_present'] or names):raise ValueError('wrong empty semantics')
        if key==('absent',None) and actual['part_present']:raise ValueError('wrong absent semantics')
        if key[1] in (1,2,3):
            if 'Decimal' in names or ('Empty' in names)!=(key[0]=='control'):raise ValueError('unexpected post-resave property')
            if len(names)!=len(set(names)):raise ValueError('duplicate names')
            control=next((p for p in actual['properties'] if p['name']=='Control'),None)
            if not control or control['value']!=('before' if key[0]=='control' else 'after'):raise ValueError('control mismatch')
    if set(paths)!=(expected|{('empty',None),('absent',None)}):raise ValueError('record set mismatch')
    for kind in ('control','edited'):
        first=inspect(paths[(kind,1)])['properties']
        if any(inspect(paths[(kind,n)])['properties']!=first for n in (2,3)):raise ValueError('later application inventories changed')
        expected_values={'Control':('lpwstr','before' if kind=='control' else 'after'),'Count':('r8','-2147483648' if kind=='control' else '2147483647'),'Date':('filetime','2020-01-02T03:04:05Z' if kind=='control' else '2021-02-03T04:05:06Z'),'Flag':('bool','0' if kind=='control' else '1'),'Real':('r8','1.25' if kind=='control' else '0.125'),'PresentationFormat':('lpwstr','On-screen Show (16:9)'),'Slides':('r8','1')}
        if kind=='control':expected_values['Empty']=('lpwstr','')
        if {p['name']:(p['type'],p['value']) for p in first}!=expected_values:raise ValueError('unexpected application rewrite')
    before=inspect(paths[('control',0)])['properties'];after=inspect(paths[('edited',0)])['properties']
    want_before=[('Empty','lpwstr',''),('Flag','bool','false'),('Count','i4','-2147483648'),('Real','r8','1.25'),('Date','filetime','2020-01-02T03:04:05.000Z'),('Decimal','decimal','9007199254740993.0001'),('Control','lpwstr','before')]
    want_after=[('Flag','bool','true'),('Count','i4','2147483647'),('Real','r8','0.125'),('Date','filetime','2021-02-03T04:05:06.000Z'),('Decimal','decimal','9007199254740994.0001'),('Control','lpwstr','after')]
    for values,want in ((before,want_before),(after,want_after)):
        if [(x['name'],x['type'],x['value']) for x in values]!=want or [x['pid'] for x in values]!=[str(i+2) for i in range(len(values))] or any(x['fmtid']!=GUID for x in values):raise ValueError('fresh authoring mismatch')
    print('Verified ten package inventories, identities, empty/absent semantics and six application resaves.')
if __name__=='__main__':main()
