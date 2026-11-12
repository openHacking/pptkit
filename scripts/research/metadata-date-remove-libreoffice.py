#!/usr/bin/env python3
"""Observe date removal and a typed offset-date control in LibreOfficeDev."""
import hashlib,json,platform,subprocess,tempfile,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];SOFFICE=Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice');BASE=ROOT/'.tmp/metadata-date-remove-libreoffice';DC='http://purl.org/dc/elements/1.1/';CP='http://schemas.openxmlformats.org/package/2006/metadata/core-properties';DT='http://purl.org/dc/terms/'
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def inventory(path):
 with zipfile.ZipFile(path) as z:
  if z.testzip() is not None or len(z.namelist())!=len(set(z.namelist())):raise ValueError('invalid ZIP')
  root=ET.fromstring(z.read('docProps/core.xml'));result={}
  for key,ns,tag in [('title',DC,'title'),('subject',DC,'subject'),('revision',CP,'revision'),('created',DT,'created'),('modified',DT,'modified')]:
   nodes=root.findall(f'{{{ns}}}{tag}')
   if len(nodes)>1:raise ValueError('duplicate property')
   result[key]=''.join(nodes[0].itertext()) if nodes else None
  return result

def main():
 BASE.mkdir(parents=True,exist_ok=True);run=Path(tempfile.mkdtemp(prefix='run-',dir=BASE));inputs=run/'input';inputs.mkdir();seed=inputs/'seed.pptx';control=inputs/'control.pptx';edited=inputs/'edited.pptx'
 generate="""import {writeFile} from 'node:fs/promises';import {createPresentation} from './packages/core/dist/index.js';import {generatePptx} from './packages/pptx-exporter/dist/index.js';const d=createPresentation({metadata:{title:'Date removal control',subject:'Sentinel',revision:7,createdAt:'2020-02-29T02:00:00.120Z',modifiedAt:'2021-01-02T03:04:05.123Z'}});d.addSlide({id:'slide'}).addElement({type:'text',content:'Date removal control',box:{x:20,y:20,width:400,height:80}});await writeFile(process.argv[1],(await generatePptx(d)).bytes);"""
 subprocess.run(['node','--input-type=module','-e',generate,str(seed)],cwd=ROOT,check=True)
 with zipfile.ZipFile(seed) as a,zipfile.ZipFile(control,'w') as b:
  for info in a.infolist():
   payload=a.read(info.filename)
   if info.filename=='docProps/core.xml':
    old=b'2020-02-29T02:00:00.120Z';new=b'2020-02-29T10:00:00.1200+08:00'
    if payload.count(old)!=1:raise ValueError('seed date mismatch')
    payload=payload.replace(old,new)
   b.writestr(info,payload)
 js="""import {readFile,writeFile} from 'node:fs/promises';import {importPptx} from './packages/pptx-parser/dist/index.js';const s=await importPptx(await readFile(process.argv[1]),{projection:'content',metadataSaveProfile:'metadata-date-remove-v14'});try{s.document.setMetadata({...s.document.metadata,createdAt:undefined,modifiedAt:undefined});const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));await writeFile(process.argv[2],(await s.save(p.plan)).bytes)}finally{s.dispose()}"""
 subprocess.run(['node','--input-type=module','-e',js,str(control),str(edited)],cwd=ROOT,check=True)
 with zipfile.ZipFile(control) as a,zipfile.ZipFile(edited) as b:
  if set(a.namelist())!=set(b.namelist()):raise ValueError('entry sets differ')
  changed=sorted(n for n in a.namelist() if a.read(n)!=b.read(n))
 if changed!=['docProps/core.xml']:raise ValueError('unexpected source edit')
 records=[];logs=[]
 for n in range(4):
  for kind in ('control','edited'):
   parent=inputs if n==0 else run/f'round-{n}';path=parent/f'{kind}.pptx'
   if n:
    parent.mkdir(exist_ok=True);previous=(inputs if n==1 else run/f'round-{n-1}')/f'{kind}.pptx'
    result=subprocess.run([str(SOFFICE),f'-env:UserInstallation={(run/f"profile-{n}-{kind}").as_uri()}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(parent),str(previous)],capture_output=True,text=True)
    logs.append({'round':n,'kind':kind,'exitCode':result.returncode,'stdout':result.stdout,'stderr':result.stderr})
    if result.returncode or not path.is_file():raise ValueError(f'conversion failed: {result.stderr}')
   records.append({'round':n,'kind':kind,'path':str(path.relative_to(ROOT)),'sha256':sha(path),'metadata':inventory(path)})
 evidence={'scope':'One offset-date control and PPTKit date removal, three LibreOfficeDev headless resaves per branch; no UI or PowerPoint.','application':subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(),'platform':platform.platform(),'runnerSha256':sha(Path(__file__)),'seed':{'path':str(seed.relative_to(ROOT)),'sha256':sha(seed)},'sourceChangedParts':changed,'records':records,'logs':logs};(ROOT/'docs/testing/metadata-date-remove-libreoffice-evidence.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2)+'\n');print(json.dumps(records,ensure_ascii=False))
if __name__=='__main__':main()
