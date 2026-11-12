#!/usr/bin/env python3
"""Verify exact OPC prefix experiments and retained headless outputs."""
import hashlib,json,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];data=json.loads((ROOT/'docs/testing/opc-prefix-libreoffice-evidence.json').read_text());previous=ROOT/'docs/testing/combined-property-create-libreoffice-evidence.json';prior=json.loads(previous.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def require(ok,message):
 if not ok:raise ValueError(message)
require(sha(ROOT/'scripts/research/opc-prefix-libreoffice.py')==data['runnerSha256'] and sha(previous)==data['seedEvidenceSha256'],'dependency hash mismatch')
seeds={x['kind']:x for x in prior['records'] if x['round']==0};require(set(seeds)=={'control','edited'},'seed count')
variants={'baseline','rel-prefix','types-prefix','both-prefix','declaration'};seen=set()
def tree(node):return (node.tag,tuple(sorted(node.attrib.items())),node.text,tuple(tree(child) for child in node))
def prefix(bytes_,ns,p,names):
 text=bytes_.decode('utf-8').replace(f'xmlns="{ns}"',f'xmlns:{p}="{ns}"',1)
 for n in names:text=text.replace(f'<{n}',f'<{p}:{n}').replace(f'</{n}>',f'</{p}:{n}>')
 return text.encode('utf-8')
for record in data['records']:
 kind=record['kind'];variant=record['variant'];key=(kind,variant);path=(ROOT/record['path']).resolve();require(key not in seen and kind in seeds and variant in variants and path.is_relative_to((ROOT/'.tmp/opc-prefix-libreoffice').resolve()) and sha(path)==record['sha256'],'input identity mismatch');seen.add(key)
 seed=ROOT/seeds[kind]['path'];require(sha(seed)==seeds[kind]['sha256'],'seed hash mismatch')
 with zipfile.ZipFile(seed) as a,zipfile.ZipFile(path) as b:
  require(b.testzip() is None and set(a.namelist())==set(b.namelist()) and len(b.namelist())==len(set(b.namelist())),'ZIP mismatch');changes=sorted(n for n in a.namelist() if a.read(n)!=b.read(n));require(changes==record['changedParts'],'changed entries mismatch')
  for name in ('_rels/.rels','[Content_Types].xml'):
   raw=a.read(name);expected=raw
   if name=='_rels/.rels' and variant in ('rel-prefix','both-prefix'):expected=prefix(raw,'http://schemas.openxmlformats.org/package/2006/relationships','r',['Relationships','Relationship'])
   if name=='[Content_Types].xml' and variant in ('types-prefix','both-prefix'):expected=prefix(raw,'http://schemas.openxmlformats.org/package/2006/content-types','t',['Types','Default','Override'])
   if variant=='declaration':require(raw.count(b' standalone="yes"')==1,'declaration absent');expected=raw.replace(b' standalone="yes"',b'',1)
   require(b.read(name)==expected and tree(ET.fromstring(raw))==tree(ET.fromstring(expected)),'declaration mismatch')
  expectedChanges=[] if variant=='baseline' else ['_rels/.rels'] if variant=='rel-prefix' else ['[Content_Types].xml'] if variant=='types-prefix' else ['[Content_Types].xml','_rels/.rels']
  require(changes==expectedChanges,'unexpected source changes')
 success=variant in ('baseline','declaration');require((record['exitCode']==0 and record['outputExists'])==success,'unexpected conversion status');output=path.parent/'output/input.pptx'
 if success:
  require(output.is_file() and sha(output)==record['outputSha256'],'output mismatch')
  with zipfile.ZipFile(output) as z:require(z.testzip() is None,'invalid output ZIP')
 else:require(not output.exists() and record['outputSha256'] is None and 'source file could not be loaded' in record['stderr'],'failure evidence mismatch')
require(seen=={(k,v) for k in ('control','edited') for v in variants},'missing variants')
print('Verified 10 exact input variants and 4 retained successful outputs.')
