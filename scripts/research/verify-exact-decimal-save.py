#!/usr/bin/env python3
"""Independently check retained exact-decimal experiment packages and manifests."""
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
BASE = (ROOT / '.tmp/exact-decimal-save-libreoffice').resolve()
CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def inventory(path):
    properties, decimals = [], []
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or len(z.namelist()) != len(set(z.namelist())): raise ValueError('invalid ZIP')
        for name in z.namelist():
            if not (name.endswith('.xml') or name.endswith('.rels')): continue
            root = ET.fromstring(z.read(name))
            decimals += [{'part':name,'value':node.text} for node in root.iter(f'{{{VT}}}decimal')]
            if root.tag == f'{{{CUSTOM}}}Properties':
                for item in root.findall(f'{{{CUSTOM}}}property'):
                    children=list(item)
                    properties.append({'name':item.get('name'),'type':children[0].tag.split('}')[-1] if len(children)==1 else None,'value':children[0].text if len(children)==1 else None})
    return sorted(properties,key=lambda item:item['name'] or ''),decimals

def main():
    evidence=Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'docs/testing/exact-decimal-save-libreoffice-evidence.json'
    data=json.loads(evidence.read_text());records=data['records'];logs=data['logs']; expected={(kind,n) for kind in ('control','edited') for n in range(4)}
    if len(records)!=8 or len(logs)!=6 or {(x['kind'],x['round']) for x in logs}!={(kind,n) for kind in ('control','edited') for n in (1,2,3)} or any(x['exit_code']!=0 for x in logs): raise ValueError('incomplete execution')
    source=ROOT/'tests/fixtures/custom-exact-decimal-source.pptx'
    if digest(source)!=data['source_sha256'] or data['pptkit_changed_parts']!=['docProps/custom.xml']: raise ValueError('source mismatch')
    paths={}
    for record in records:
        key=(record['kind'],record['round']);file=(ROOT/record['path']).resolve()
        if key not in expected or key in paths or not file.is_relative_to(BASE) or digest(file)!=record['sha256']: raise ValueError('record hash or location mismatch')
        properties,decimals=inventory(file)
        if properties!=record['properties'] or decimals!=record['decimal_elements']: raise ValueError('inventory mismatch')
        values=['0.1','9007199254740993.0001','-0.0001'] if key==('control',0) else ['0.2','9007199254740994.0001','-0.0001'] if key==('edited',0) else []
        if [item['value'] for item in decimals]!=values or not any(item['name']=='Control' and item['value']=='before' for item in properties) or key[1] > 0 and any(item['name'] in {'Tenth','Wide','Negative'} for item in properties): raise ValueError('unexpected observed values')
        paths[key]=file
    if set(paths)!=expected or digest(paths[('control',0)])!=data['source_sha256'] or digest(paths[('edited',0)])!=data['pptkit_edited_sha256']: raise ValueError('baseline mismatch')
    with zipfile.ZipFile(paths[('control',0)]) as a, zipfile.ZipFile(paths[('edited',0)]) as b:
        if set(a.namelist())!=set(b.namelist()) or sorted(n for n in a.namelist() if a.read(n)!=b.read(n))!=['docProps/custom.xml']: raise ValueError('unexpected source delta')
        original=a.read('docProps/custom.xml'); expected_xml=original
        for old,new in [(b'0.1',b'0.2'),(b'9007199254740993.0001',b'9007199254740994.0001')]:
            marker=b'<vt:decimal>'+old+b'</vt:decimal>'
            if expected_xml.count(marker)!=1: raise ValueError('ambiguous source text')
            expected_xml=expected_xml.replace(marker,b'<vt:decimal>'+new+b'</vt:decimal>',1)
        if expected_xml!=b.read('docProps/custom.xml'): raise ValueError('unexpected XML edit')
    print('Verified eight exact-decimal inventories, hashes, source changes and six headless resaves.')
if __name__=='__main__': main()
