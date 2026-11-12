#!/usr/bin/env python3
"""Independently verify retained three-property creation evidence and package contents."""
import hashlib, json, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]
manifest=Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/'docs/testing/extended-create-libreoffice-evidence.json'
data=json.loads(manifest.read_text())
def require(condition,message):
    if not condition: raise ValueError(message)
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
require(sha(ROOT/'scripts/research/extended-create-libreoffice.py')==data['runnerSha256'],'runner mismatch')
seed=data['seed']; require(sha(ROOT/seed['path'])==seed['sha256'],'seed mismatch')
expected={(round,kind) for round in range(4) for kind in ('control','edited')}
require(len(data['records'])==8 and len(data['logs'])==6 and {(x['round'],x['kind']) for x in data['logs']}=={(n,k) for n in (1,2,3) for k in ('control','edited')} and all(x['status']=='saved' and x['exitCode']==0 for x in data['logs']),'incomplete conversions')
core='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'; base='http://schemas.openxmlformats.org/officeDocument/2006/relationships/'; ep='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'; dc='http://purl.org/dc/elements/1.1/'
ct='http://schemas.openxmlformats.org/package/2006/content-types'; custom_ns='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'; core_ns='http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
def read(path):
    with zipfile.ZipFile(path) as z:
        require(z.testzip() is None and len(z.namelist())==len(set(z.namelist())),'invalid ZIP')
        rels=ET.fromstring(z.read('_rels/.rels')); types=ET.fromstring(z.read('[Content_Types].xml')); result={}
        for kind,rel in [('core',core),('extended',base+'extended-properties'),('custom',base+'custom-properties')]:
            matches=[r for r in rels if r.get('Type')==rel]; require(len(matches)<=1,'ambiguous relationship')
            if not matches: result[kind]=None; continue
            require(matches[0].get('TargetMode','Internal')=='Internal','external metadata target')
            target=matches[0].get('Target','').lstrip('/'); require(target and '..' not in target.split('/') and target in z.namelist(),'invalid target'); root=ET.fromstring(z.read(target))
            mime='application/vnd.openxmlformats-package.core-properties+xml' if kind=='core' else f'application/vnd.openxmlformats-officedocument.{kind}-properties+xml'
            overrides=[node for node in types if node.tag==f'{{{ct}}}Override' and node.get('PartName')=='/'+target]
            require(len(overrides)==1 and overrides[0].get('ContentType')==mime,'metadata content type mismatch')
            expected_root={'core':f'{{{core_ns}}}coreProperties','extended':f'{{{ep}}}Properties','custom':f'{{{custom_ns}}}Properties'}[kind]
            require(root.tag==expected_root,'metadata root mismatch')
            if kind=='custom':
                require(all(len(node)==1 for node in root),'invalid custom')
                result[kind]=[{'name':n.get('name'),'type':n[0].tag.split('}')[-1],'value':''.join(n[0].itertext())} for n in root]
            else:
                tag,ns=('title',dc) if kind=='core' else ('Company',ep); nodes=root.findall(f'{{{ns}}}{tag}'); require(len(nodes)<=1,'duplicate value'); result[kind]=''.join(nodes[0].itertext()) if nodes else None
        return result
paths={}
for record in data['records']:
    key=(record['round'],record['kind']); path=(ROOT/record['path']).resolve()
    require(key in expected and key not in paths and path.is_relative_to((ROOT/'.tmp/extended-create-libreoffice').resolve()) and sha(path)==record['sha256'],'record identity mismatch')
    actual=read(path); require(actual==record['metadata'],'inventory mismatch')
    if key==(0,'control'): require(actual=={'core':None,'extended':None,'custom':None},'control mismatch')
    if key[1]=='edited':
        require(actual['core']=='Created 中文' and actual['extended']=='Company & 中文' and actual['custom']==[{'name':'Count','type':'i4' if key[0]==0 else 'r8','value':'42'}],'edited values mismatch')
    paths[key]=path
require(set(paths)==expected,'missing record')
with zipfile.ZipFile(ROOT/seed['path']) as a,zipfile.ZipFile(paths[(0,'control')]) as b:
    require(set(a.namelist())-set(b.namelist())=={'docProps/core.xml','docProps/app.xml'} and not set(b.namelist())-set(a.namelist()) and sorted(name for name in b.namelist() if a.read(name)!=b.read(name))==['[Content_Types].xml','_rels/.rels'],'control construction mismatch')
    for name in ['_rels/.rels','[Content_Types].xml']:
        before=ET.fromstring(a.read(name)); after=ET.fromstring(b.read(name))
        removed=[node for node in before if node.get('Target') in ('docProps/core.xml','docProps/app.xml') or node.get('PartName') in ('/docProps/core.xml','/docProps/app.xml')]
        require(len(removed)==2,'missing seed declarations')
        for node in removed: before.remove(node)
        def tree(node): return (node.tag,tuple(sorted(node.attrib.items())),node.text,tuple(tree(child) for child in node))
        require(tree(before)==tree(after),'unexpected declaration change')
with zipfile.ZipFile(paths[(0,'control')]) as a,zipfile.ZipFile(paths[(0,'edited')]) as b:
    added=sorted(set(b.namelist())-set(a.namelist())); changed=sorted(name for name in a.namelist() if a.read(name)!=b.read(name))
    require(not set(a.namelist())-set(b.namelist()) and added==data['sourceAddedParts']==['pptkit-core-properties-1.xml','pptkit-custom-properties-1.xml','pptkit-extended-properties-1.xml'] and changed==data['sourceChangedParts']==['[Content_Types].xml','_rels/.rels'],'source isolation mismatch')
print('Verified eight actual package inventories, hashes, conversion records and source isolation.')
