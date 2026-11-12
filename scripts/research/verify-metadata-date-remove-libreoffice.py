#!/usr/bin/env python3
"""Independently verify source-date removal and retained LibreOffice observations."""
import hashlib,json,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];manifest=Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'docs/testing/metadata-date-remove-libreoffice-evidence.json';data=json.loads(manifest.read_text())
def require(ok,message):
 if not ok:raise ValueError(message)
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
require(sha(ROOT/'scripts/research/metadata-date-remove-libreoffice.py')==data['runnerSha256'],'runner mismatch');seed=ROOT/data['seed']['path'];require(sha(seed)==data['seed']['sha256'],'seed mismatch');expected={(n,k) for n in range(4) for k in ('control','edited')};require(len(data['records'])==8 and len(data['logs'])==6 and {(r['round'],r['kind']) for r in data['logs']}=={(n,k) for n in (1,2,3) for k in ('control','edited')} and all(r['exitCode']==0 for r in data['logs']),'incomplete conversions')
dc='http://purl.org/dc/elements/1.1/';dt='http://purl.org/dc/terms/';cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties';rel='http://schemas.openxmlformats.org/package/2006/relationships';ct='http://schemas.openxmlformats.org/package/2006/content-types'
def inventory(path):
 with zipfile.ZipFile(path) as z:
  require(z.testzip() is None and len(z.namelist())==len(set(z.namelist())),'invalid ZIP');rels=ET.fromstring(z.read('_rels/.rels'));matches=[x for x in rels if x.tag==f'{{{rel}}}Relationship' and x.get('Type')==rel+'/metadata/core-properties'];require(len(matches)==1 and matches[0].get('Target','').lstrip('/')=='docProps/core.xml' and matches[0].get('TargetMode','Internal')=='Internal','invalid relationship');types=ET.fromstring(z.read('[Content_Types].xml'));overrides=[x for x in types if x.tag==f'{{{ct}}}Override' and x.get('PartName')=='/docProps/core.xml'];require(len(overrides)==1 and overrides[0].get('ContentType')=='application/vnd.openxmlformats-package.core-properties+xml','invalid MIME');root=ET.fromstring(z.read('docProps/core.xml'));require(root.tag==f'{{{cp}}}coreProperties','invalid root');result={}
  for key,ns,name in [('title',dc,'title'),('subject',dc,'subject'),('revision',cp,'revision'),('created',dt,'created'),('modified',dt,'modified')]:
   nodes=root.findall(f'{{{ns}}}{name}');require(len(nodes)<=1,'duplicate property');result[key]=''.join(nodes[0].itertext()) if nodes else None
  return result
paths={}
for record in data['records']:
 key=(record['round'],record['kind']);path=(ROOT/record['path']).resolve();require(key in expected and key not in paths and path.is_relative_to((ROOT/'.tmp/metadata-date-remove-libreoffice').resolve()) and sha(path)==record['sha256'],'record identity mismatch');actual=inventory(path);require(actual==record['metadata'] and (actual['title'],actual['subject'],actual['revision'])==('Date removal control','Sentinel','7'),'inventory mismatch')
 dates=(actual['created'],actual['modified']);target=(None,None) if key[1]=='edited' else ('2020-02-29T10:00:00.1200+08:00','2021-01-02T03:04:05.123Z') if key[0]==0 else ('2020-02-29T02:00:00Z','2021-01-02T03:04:05Z');require(dates==target,'date observation mismatch');paths[key]=path
require(set(paths)==expected,'missing records')
with zipfile.ZipFile(seed) as a,zipfile.ZipFile(paths[(0,'control')]) as b:
 require(set(a.namelist())==set(b.namelist()) and sorted(n for n in a.namelist() if a.read(n)!=b.read(n))==['docProps/core.xml'],'control derivation mismatch');old=b'2020-02-29T02:00:00.120Z';require(a.read('docProps/core.xml').count(old)==1 and a.read('docProps/core.xml').replace(old,b'2020-02-29T10:00:00.1200+08:00')==b.read('docProps/core.xml'),'offset control mismatch')
with zipfile.ZipFile(paths[(0,'control')]) as a,zipfile.ZipFile(paths[(0,'edited')]) as b:
 require(set(a.namelist())==set(b.namelist()) and sorted(n for n in a.namelist() if a.read(n)!=b.read(n))==data['sourceChangedParts']==['docProps/core.xml'],'source delta mismatch');original=a.read('docProps/core.xml')
 for target in [b'<cpTerms:created xsi:type="cpTerms:W3CDTF">2020-02-29T10:00:00.1200+08:00</cpTerms:created>',b'<cpTerms:modified xsi:type="cpTerms:W3CDTF">2021-01-02T03:04:05.123Z</cpTerms:modified>']:
  require(original.count(target)==1,'source target mismatch');original=original.replace(target,b'')
 require(original==b.read('docProps/core.xml'),'non-target bytes changed')
print('Verified eight packages, offset control derivation, exact deletion spans and three-round date observations.')
