#!/usr/bin/env python3
"""Independently verify retained paired date edit/insertion observations."""
import hashlib,json,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];data=json.loads((ROOT/'docs/testing/core-date-libreoffice-evidence.json').read_text())
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
def require(condition,message):
 if not condition: raise ValueError(message)
require(sha(ROOT/'scripts/research/core-date-libreoffice.py')==data['runnerSha256'],'runner mismatch')
require(data['sourceChangedParts']==['docProps/core.xml'],'wrong source delta')
require(len(data['records'])==8 and len(data['logs'])==6 and all(x['exitCode']==0 for x in data['logs']),'incomplete process records')
ns={'title':'http://purl.org/dc/elements/1.1/','creator':'http://purl.org/dc/elements/1.1/','subject':'http://purl.org/dc/elements/1.1/','description':'http://purl.org/dc/elements/1.1/','language':'http://purl.org/dc/elements/1.1/','revision':'http://schemas.openxmlformats.org/package/2006/metadata/core-properties','keywords':'http://schemas.openxmlformats.org/package/2006/metadata/core-properties','created':'http://purl.org/dc/terms/','modified':'http://purl.org/dc/terms/'}
paths={}
for record in data['records']:
 key=(record['round'],record['kind']);path=(ROOT/record['path']).resolve()
 require(key not in paths and key in {(i,k) for i in range(4) for k in ('control','edited')} and path.is_relative_to((ROOT/'.tmp/core-date-libreoffice').resolve()) and sha(path)==record['sha256'],'record mismatch')
 with zipfile.ZipFile(path) as z:
  require(z.testzip() is None and len(z.namelist())==len(set(z.namelist())),'invalid ZIP')
  root=ET.fromstring(z.read('docProps/core.xml'));actual={}
  for name,uri in ns.items():
   elements=root.findall(f'{{{uri}}}{name}');require(len(elements)<=1,'duplicate property');actual[name]=''.join(elements[0].itertext()) if elements else None
   if elements and name in ('created','modified'):
    prefix = ('dt' if name=='created' else 'dct') if key[0]==0 else 'dcterms'
    require(elements[0].get('{http://www.w3.org/2001/XMLSchema-instance}type') == f'{prefix}:W3CDTF','date type mismatch')
  require(actual==record['metadata'],'inventory mismatch')
 require(actual['creator']=='Sentinel Author' and actual['revision']=='7','control properties changed')
 require(actual['keywords']==('alpha, beta' if key[0]==0 else 'alpha beta'),'keyword mismatch')
 if key==(0,'control'): dates=['2020-01-02T11:04:05.1200+08:00',None]
 elif key==(0,'edited'): dates=['2022-03-04T05:06:07.123Z','2023-04-05T06:07:08.456Z']
 elif key[1]=='control': dates=['2020-01-02T03:04:05Z',None]
 else: dates=['2022-03-04T05:06:07Z','2023-04-05T06:07:08Z']
 require([actual['created'],actual['modified']]==dates,'unexpected dates');paths[key]=path
require(len(paths)==8,'missing rounds')
with zipfile.ZipFile(paths[(0,'control')]) as a,zipfile.ZipFile(paths[(0,'edited')]) as b:
 require(set(a.namelist())==set(b.namelist()) and sorted(n for n in a.namelist() if a.read(n)!=b.read(n))==['docProps/core.xml'],'package delta')
 old=a.read('docProps/core.xml');new=b.read('docProps/core.xml')
 before=b'2020-01-02T11:04:05.1200+08:00';after=b'2022-03-04T05:06:07.123Z';end=b'</ns0:coreProperties>'
 insertion=b'<dct:modified xmlns:dct="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dct:W3CDTF">2023-04-05T06:07:08.456Z</dct:modified>'
 require(old.count(before)==1 and old.count(end)==1 and old.replace(before,after).replace(end,insertion+end)==new,'unexpected source edits')
print('Verified eight date inventories, source byte isolation and six headless resaves.')
