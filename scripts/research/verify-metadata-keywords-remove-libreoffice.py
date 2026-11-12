#!/usr/bin/env python3
"""Independently verify keyword removal and retained LibreOffice observations."""
import hashlib,json,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];manifest=Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'docs/testing/metadata-keywords-remove-libreoffice-evidence.json';data=json.loads(manifest.read_text())
def require(value,message):
 if not value:raise ValueError(message)
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
require(sha(ROOT/'scripts/research/metadata-keywords-remove-libreoffice.py')==data['runnerSha256'],'runner mismatch');seed=ROOT/data['seed']['path'];require(sha(seed)==data['seed']['sha256'],'seed mismatch');expected={(n,k) for n in range(4) for k in ('control','edited')};require(len(data['records'])==8 and len(data['logs'])==6 and {(r['round'],r['kind']) for r in data['logs']}=={(n,k) for n in (1,2,3) for k in ('control','edited')} and all(r['exitCode']==0 for r in data['logs']),'incomplete conversions')
dc='http://purl.org/dc/elements/1.1/';cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties';rel='http://schemas.openxmlformats.org/package/2006/relationships';ct='http://schemas.openxmlformats.org/package/2006/content-types'
def inventory(path):
 with zipfile.ZipFile(path) as z:
  require(z.testzip() is None and len(z.namelist())==len(set(z.namelist())),'invalid ZIP');rootrels=ET.fromstring(z.read('_rels/.rels'));matches=[x for x in rootrels if x.tag==f'{{{rel}}}Relationship' and x.get('Type')==rel+'/metadata/core-properties'];require(len(matches)==1 and matches[0].get('Target','').lstrip('/')=='docProps/core.xml' and matches[0].get('TargetMode','Internal')=='Internal','invalid relationship');types=ET.fromstring(z.read('[Content_Types].xml'));overrides=[x for x in types if x.tag==f'{{{ct}}}Override' and x.get('PartName')=='/docProps/core.xml'];require(len(overrides)==1 and overrides[0].get('ContentType')=='application/vnd.openxmlformats-package.core-properties+xml','invalid MIME');root=ET.fromstring(z.read('docProps/core.xml'));require(root.tag==f'{{{cp}}}coreProperties','invalid root');result={}
  for key,ns in [('title',dc),('creator',dc),('revision',cp),('keywords',cp)]:
   nodes=root.findall(f'{{{ns}}}{key}');require(len(nodes)<=1,'duplicate property');result[key]=''.join(nodes[0].itertext()) if nodes else None
  return result
paths={}
for record in data['records']:
 key=(record['round'],record['kind']);path=(ROOT/record['path']).resolve();require(key in expected and key not in paths and path.is_relative_to((ROOT/'.tmp/metadata-keywords-remove-libreoffice').resolve()) and sha(path)==record['sha256'],'record identity mismatch');actual=inventory(path);require(actual==record['metadata'] and (actual['title'],actual['creator'],actual['revision'])==('Original title','Sentinel Author','7'),'inventory mismatch');expected_keyword=None if key[1]=='edited' else 'alpha, beta' if key[0]==0 else 'alpha beta';require(actual['keywords']==expected_keyword,'keyword observation mismatch');paths[key]=path
require(set(paths)==expected and sha(paths[(0,'control')])==sha(seed),'missing records or control mismatch')
with zipfile.ZipFile(paths[(0,'control')]) as a,zipfile.ZipFile(paths[(0,'edited')]) as b:
 require(set(a.namelist())==set(b.namelist()) and sorted(n for n in a.namelist() if a.read(n)!=b.read(n))==data['sourceChangedParts']==['docProps/core.xml'],'source delta mismatch');source=a.read('docProps/core.xml');target=b'<cp:keywords>alpha, beta</cp:keywords>';require(source.count(target)==1 and source.replace(target,b'')==b.read('docProps/core.xml'),'non-target bytes changed')
print('Verified eight packages, exact keyword deletion span and three-round observations.')
