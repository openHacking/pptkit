#!/usr/bin/env python3
"""Independently check retained source removal and LibreOffice XML observations."""
import hashlib,json,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]
manifest=Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'docs/testing/metadata-remove-libreoffice-evidence.json';data=json.loads(manifest.read_text())
def require(value,message):
 if not value:raise ValueError(message)
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
require(sha(ROOT/'scripts/research/metadata-remove-libreoffice.py')==data['runnerSha256'],'runner mismatch')
expected={(n,k) for n in range(4) for k in ('control','edited')}
require(len(data['records'])==8 and len(data['logs'])==6 and {(x['round'],x['kind']) for x in data['logs']}=={(n,k) for n in (1,2,3) for k in ('control','edited')} and all(x['exitCode']==0 for x in data['logs']),'incomplete conversions')
dc='http://purl.org/dc/elements/1.1/';cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties';ep='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';rels_ns='http://schemas.openxmlformats.org/package/2006/relationships';ct='http://schemas.openxmlformats.org/package/2006/content-types'
def inventory(path):
 with zipfile.ZipFile(path) as z:
  require(z.testzip() is None and len(z.namelist())==len(set(z.namelist())),'invalid ZIP');rels=ET.fromstring(z.read('_rels/.rels'));types=ET.fromstring(z.read('[Content_Types].xml'))
  for part,typ,mime in [('docProps/core.xml','http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties','application/vnd.openxmlformats-package.core-properties+xml'),('docProps/app.xml','http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties','application/vnd.openxmlformats-officedocument.extended-properties+xml')]:
   matches=[r for r in rels if r.tag==f'{{{rels_ns}}}Relationship' and r.get('Type')==typ];require(len(matches)==1 and matches[0].get('Target','').lstrip('/')==part and matches[0].get('TargetMode','Internal')=='Internal','invalid metadata relationship');overrides=[n for n in types if n.tag==f'{{{ct}}}Override' and n.get('PartName')=='/'+part];require(len(overrides)==1 and overrides[0].get('ContentType')==mime,'invalid metadata MIME')
  core=ET.fromstring(z.read('docProps/core.xml'));app=ET.fromstring(z.read('docProps/app.xml'));require(core.tag==f'{{{cp}}}coreProperties' and app.tag==f'{{{ep}}}Properties','invalid metadata root');result={}
  for name,root,ns,local in [('title',core,dc,'title'),('subject',core,dc,'subject'),('description',core,dc,'description'),('revision',core,cp,'revision'),('company',app,ep,'Company'),('application',app,ep,'Application')]:
   found=root.findall(f'{{{ns}}}{local}');require(len(found)<=1,'duplicate property');result[name]=''.join(found[0].itertext()) if found else None
  return result
paths={}
for record in data['records']:
 key=(record['round'],record['kind']);path=(ROOT/record['path']).resolve();require(key in expected and key not in paths and path.is_relative_to((ROOT/'.tmp/metadata-remove-libreoffice').resolve()) and sha(path)==record['sha256'],'record identity mismatch');actual=inventory(path);require(actual==record['metadata'],'inventory mismatch');require(actual['title']=='Removal control' and actual['revision']=='7','sentinel mismatch')
 if key[1]=='control':require((actual['subject'],actual['description'],actual['company'])==('Subject 中文','Description & control','Company 中文'),'control values changed')
 else:require((actual['subject'],actual['description'],actual['company'])==((None,None,None) if key[0]==0 else ('','',None)),'removal observation mismatch')
 paths[key]=path
require(set(paths)==expected,'missing records')
with zipfile.ZipFile(paths[(0,'control')]) as a,zipfile.ZipFile(paths[(0,'edited')]) as b:
 require(set(a.namelist())==set(b.namelist()) and sorted(n for n in a.namelist() if a.read(n)!=b.read(n))==data['sourceChangedParts']==['docProps/app.xml','docProps/core.xml'],'source package delta mismatch')
 for name,markers in [('docProps/core.xml',['<dc:subject>Subject 中文</dc:subject>','<dc:description>Description &amp; control</dc:description>']),('docProps/app.xml',['<Company>Company 中文</Company>'])]:
  original=a.read(name)
  for marker in markers:
   target=marker.encode();require(original.count(target)==1,'source marker mismatch');original=original.replace(target,b'')
  require(original==b.read(name),'non-target XML bytes changed')
print('Verified eight actual packages, exact removal spans and three-round absence versus empty values.')
