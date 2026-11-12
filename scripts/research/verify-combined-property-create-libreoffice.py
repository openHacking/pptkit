#!/usr/bin/env python3
"""Independent ZIP/XML verification of the combined property-creation experiment."""
import hashlib,json,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];data=json.loads((ROOT/'docs/testing/combined-property-create-libreoffice-evidence.json').read_text())
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
def require(value,message):
 if not value:raise ValueError(message)
require(sha(ROOT/'scripts/research/combined-property-create-libreoffice.py')==data['runnerSha256'],'runner mismatch')
seed=json.loads((ROOT/'docs/testing/core-metadata-libreoffice-evidence.json').read_text())['generated'];require(seed['sha256']==data['seedSha256'] and sha(ROOT/seed['path'])==data['seedSha256'],'seed mismatch')
require(data['sourceAddedParts']==['pptkit-core-properties-1.xml','pptkit-custom-properties-1.xml'] and data['sourceChangedParts']==['[Content_Types].xml','_rels/.rels'],'source delta mismatch')
require(len(data['records'])==8 and len(data['logs'])==6 and {(x['round'],x['kind']) for x in data['logs']}=={(i,k) for i in (1,2,3) for k in ('control','edited')} and all(x['status']=='saved' and x['exitCode']==0 for x in data['logs']),'incomplete rounds')
coreRel='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';customRel='http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';DC='http://purl.org/dc/elements/1.1/';TERMS='http://purl.org/dc/terms/'
def inventory(path):
 with zipfile.ZipFile(path) as z:
  require(z.testzip() is None and len(z.namelist())==len(set(z.namelist())),'invalid ZIP');rels=ET.fromstring(z.read('_rels/.rels'));result={}
  for key,relation in [('core',coreRel),('custom',customRel)]:
   matches=[r for r in rels if r.get('Type')==relation];require(len(matches)<=1,'ambiguous relationship')
   if not matches:result[key]={} if key=='core' else [];continue
   target=matches[0].get('Target','').lstrip('/');require(target and '..' not in target.split('/') and target in z.namelist(),'invalid target');root=ET.fromstring(z.read(target))
   if key=='core':
    result[key]={}
    for name,ns in [('title',DC),('creator',DC),('created',TERMS)]:
     nodes=root.findall(f'{{{ns}}}{name}');require(len(nodes)<=1,'duplicate core property');result[key][name]=''.join(nodes[0].itertext()) if nodes else None
   else:
    require(all(len(node)==1 for node in root),'invalid custom property');result[key]=[{'name':p.get('name'),'pid':p.get('pid'),'type':p[0].tag.split('}')[-1],'value':''.join(p[0].itertext())} for p in root]
  return result
paths={}
for r in data['records']:
 key=(r['round'],r['kind']);path=(ROOT/r['path']).resolve();require(key not in paths and key in {(i,k) for i in range(4) for k in ('control','edited')} and path.is_relative_to((ROOT/'.tmp/combined-property-create-libreoffice').resolve()) and sha(path)==r['sha256'],'bad record');actual=inventory(path);require(actual==r['metadata'],'inventory mismatch')
 if key==(0,'control'):require(actual=={'core':{},'custom':[]},'control not empty')
 if key==(0,'edited'):require(actual['core']=={'title':'Created 中文','creator':'Fixture Author','created':'2020-01-02T03:04:05.123Z'} and [(p['name'],p['type'],p['value']) for p in actual['custom']]==[('Text','lpwstr','中文 & value'),('Count','i4','42')],'edited source mismatch')
 if key[0]>0 and key[1]=='edited':
  require(actual['core']=={'title':'Created 中文','creator':'Fixture Author','created':'2020-01-02T03:04:05Z'},'edited dates changed unexpectedly');props={p['name']:p for p in actual['custom']};require(props['Text']['value']=='中文 & value' and props['Count']['type']=='r8' and props['Count']['value']=='42','edited custom mismatch')
 paths[key]=path
require(len(paths)==8,'missing round')
with zipfile.ZipFile(paths[(0,'control')]) as a,zipfile.ZipFile(paths[(0,'edited')]) as b:
 require(sorted(set(b.namelist())-set(a.namelist()))==data['sourceAddedParts'] and not set(a.namelist())-set(b.namelist()) and sorted(n for n in a.namelist() if a.read(n)!=b.read(n))==data['sourceChangedParts'],'source isolation failed')
def tree(node):return (node.tag,tuple(sorted(node.attrib.items())),node.text,tuple(tree(x) for x in node))
for probe in data['prefixProbe']:
 path=(ROOT/probe['path']).resolve();require(path.is_relative_to((ROOT/'.tmp/combined-property-create-libreoffice').resolve()) and sha(path)==probe['sha256'] and probe['exitCode']!=0 and not probe['outputExists'],'prefix probe mismatch')
 source=paths[(0,probe['kind'])]
 with zipfile.ZipFile(source) as a,zipfile.ZipFile(path) as b:
  require(set(a.namelist())==set(b.namelist()) and sorted(n for n in a.namelist() if a.read(n)!=b.read(n))==['[Content_Types].xml','_rels/.rels'],'prefix probe changed other parts')
  for name in ['[Content_Types].xml','_rels/.rels']:require(tree(ET.fromstring(a.read(name)))==tree(ET.fromstring(b.read(name))),'prefix probe changed XML structure')
 require(inventory(path)==inventory(source),'prefix probe changed metadata')
require({p['kind'] for p in data['prefixProbe']}=={'control','edited'},'prefix probes missing')
print('Verified eight package inventories, source isolation and two structurally equivalent prefix controls.')
