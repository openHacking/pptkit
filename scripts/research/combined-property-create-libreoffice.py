#!/usr/bin/env python3
"""Observe simultaneous core and custom property creation under LibreOfficeDev."""
import hashlib,importlib.util,json,platform,subprocess,tempfile,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]
SOFFICE=Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
BASE=ROOT/'.tmp/combined-property-create-libreoffice'
CORE_REL='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
DC='http://purl.org/dc/elements/1.1/'; TERMS='http://purl.org/dc/terms/'
ORACLE=ROOT/'scripts/research/custom-property-create-libreoffice.py'
spec=importlib.util.spec_from_file_location('oracle',ORACLE);oracle=importlib.util.module_from_spec(spec);spec.loader.exec_module(oracle)
def digest(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def inventory(path):
 with zipfile.ZipFile(path) as z:
  if z.testzip() is not None or len(z.namelist())!=len(set(z.namelist())):raise ValueError('invalid ZIP')
  matches=[r for r in ET.fromstring(z.read('_rels/.rels')) if r.get('Type')==CORE_REL]
  if len(matches)>1:raise ValueError('ambiguous core relation')
  values={}
  if matches:
   name=matches[0].get('Target','').lstrip('/')
   if not name or '..' in name.split('/') or name not in z.namelist():raise ValueError('invalid core target')
   root=ET.fromstring(z.read(name))
   for key,ns in [('title',DC),('creator',DC),('created',TERMS)]:
    entries=root.findall(f'{{{ns}}}{key}')
    if len(entries)>1:raise ValueError('duplicate core property')
    values[key]=''.join(entries[0].itertext()) if entries else None
  return {'core':values,'custom':oracle.inventory(path)}
def main():
 prior=json.loads((ROOT/'docs/testing/core-metadata-libreoffice-evidence.json').read_text())['generated'];seed=ROOT/prior['path']
 if digest(seed)!=prior['sha256'] or oracle.inventory(seed):raise ValueError('seed mismatch')
 BASE.mkdir(parents=True,exist_ok=True);run=Path(tempfile.mkdtemp(prefix='run-',dir=BASE));inputs=run/'input';inputs.mkdir();control=inputs/'control.pptx';edited=inputs/'edited.pptx'
 with zipfile.ZipFile(seed) as a,zipfile.ZipFile(control,'w') as b:
  for info in a.infolist():
   if info.filename=='docProps/core.xml':continue
   payload=a.read(info.filename)
   if info.filename=='_rels/.rels':
    root=ET.fromstring(payload);found=[r for r in root if r.get('Type')==CORE_REL]
    if len(found)!=1 or found[0].get('Target')!='docProps/core.xml':raise ValueError('unexpected seed relation')
    marker=b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    if payload.count(marker)!=1:raise ValueError('unexpected seed relationship text')
    payload=payload.replace(marker,b'')
   elif info.filename=='[Content_Types].xml':
    root=ET.fromstring(payload);found=[e for e in root if e.get('PartName')=='/docProps/core.xml']
    if len(found)!=1:raise ValueError('unexpected seed content types')
    marker=b'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    if payload.count(marker)!=1:raise ValueError('unexpected seed content type text')
    payload=payload.replace(marker,b'')
   b.writestr(info,payload)
 if inventory(control)!={'core':{},'custom':[]}:raise ValueError('control not empty')
 js="""import {readFile,writeFile} from 'node:fs/promises';import {importPptx} from './packages/pptx-parser/dist/index.js';
const s=await importPptx(await readFile(process.argv[1]),{projection:'content',metadataSaveProfile:'core-create-v8',customPropertySaveProfile:'create-and-append-v5'});
try{s.document.setMetadata({...s.document.metadata,title:'Created 中文',author:'Fixture Author',createdAt:'2020-01-02T03:04:05.123Z',customProperties:[{name:'Text',type:'string',value:'中文 & value'},{name:'Count',type:'integer',value:42}]});const p=await s.prepareSave({mode:'edit',customPropertyAdditions:[0,1]});if(p.status!=='ready')throw Error(JSON.stringify(p.report));await writeFile(process.argv[2],(await s.save(p.plan)).bytes)}finally{s.dispose()}"""
 subprocess.run(['node','--input-type=module','-e',js,str(control),str(edited)],cwd=ROOT,check=True)
 with zipfile.ZipFile(control) as a,zipfile.ZipFile(edited) as b:
  added=sorted(set(b.namelist())-set(a.namelist()));removed=sorted(set(a.namelist())-set(b.namelist()));changed=sorted(n for n in a.namelist() if a.read(n)!=b.read(n))
 if added!=['pptkit-core-properties-1.xml','pptkit-custom-properties-1.xml'] or removed or changed!=['[Content_Types].xml','_rels/.rels']:raise ValueError('unexpected package delta')
 prefixProbe=[]
 for kind,original in [('control',control),('edited',edited)]:
  transformed=inputs/f'prefix-{kind}.pptx';out=run/f'prefix-out-{kind}';out.mkdir()
  with zipfile.ZipFile(original) as a,zipfile.ZipFile(transformed,'w') as b:
   for info in a.infolist():
    payload=a.read(info.filename)
    if info.filename in ('_rels/.rels','[Content_Types].xml'):payload=ET.tostring(ET.fromstring(payload),encoding='utf-8',xml_declaration=True)
    b.writestr(info,payload)
  with zipfile.ZipFile(original) as a,zipfile.ZipFile(transformed) as b:
   delta=sorted(n for n in a.namelist() if a.read(n)!=b.read(n))
  if delta!=['[Content_Types].xml','_rels/.rels'] or inventory(original)!=inventory(transformed):raise ValueError('prefix probe changed semantics')
  result=subprocess.run([str(SOFFICE),f'-env:UserInstallation={(run/f"prefix-profile-{kind}").as_uri()}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(out),str(transformed)],capture_output=True,text=True)
  prefixProbe.append({'kind':kind,'path':str(transformed.relative_to(ROOT)),'sha256':digest(transformed),'changedParts':delta,'exitCode':result.returncode,'outputExists':(out/transformed.name).is_file(),'stdout':result.stdout,'stderr':result.stderr})
 records=[];logs=[]
 for n in range(4):
  for kind in ('control','edited'):
   parent=inputs if n==0 else run/f'round-{n}';path=parent/f'{kind}.pptx'
   if n:
    parent.mkdir(exist_ok=True);previous=(inputs if n==1 else run/f'round-{n-1}')/f'{kind}.pptx'
    if not previous.is_file():
     logs.append({'round':n,'kind':kind,'status':'skipped','reason':'Previous output is unavailable.'});continue
    result=subprocess.run([str(SOFFICE),f'-env:UserInstallation={(run/f"profile-{n}-{kind}").as_uri()}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(parent),str(previous)],capture_output=True,text=True)
    logs.append({'round':n,'kind':kind,'status':'saved' if result.returncode==0 and path.is_file() else 'failed','exitCode':result.returncode,'stdout':result.stdout,'stderr':result.stderr})
    if result.returncode or not path.is_file():continue
   records.append({'round':n,'kind':kind,'path':str(path.relative_to(ROOT)),'sha256':digest(path),'metadata':inventory(path)})
 evidence={'scope':'One synthetic missing core/custom control and PPTKit combined creation, LibreOfficeDev headless conversion attempts over three rounds; failed branches remain recorded, no UI or PowerPoint.','application':subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(),'platform':platform.platform(),'runnerSha256':digest(Path(__file__)),'seedSha256':digest(seed),'sourceChangedParts':changed,'sourceAddedParts':added,'prefixProbe':prefixProbe,'records':records,'logs':logs}
 (ROOT/'docs/testing/combined-property-create-libreoffice-evidence.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2)+'\n')
 print(json.dumps(records,ensure_ascii=False))
if __name__=='__main__':main()
