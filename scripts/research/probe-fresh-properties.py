#!/usr/bin/env python3
"""Compare fresh public-API custom properties with headless application resaves."""
import hashlib
import json
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT=Path(__file__).resolve().parents[2]
CUSTOM='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
VT='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def inspect(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or len(z.namelist())!=len(set(z.namelist())): raise ValueError('invalid ZIP')
        relationships=[item for item in ET.fromstring(z.read('_rels/.rels')) if item.get('Type')==REL]
        if len(relationships)>1: raise ValueError('duplicate custom relationships')
        parts=[name for name in z.namelist() if name=='docProps/custom.xml']
        properties=[]
        if parts:
            root=ET.fromstring(z.read(parts[0]));
            if root.tag!=f'{{{CUSTOM}}}Properties': raise ValueError('wrong custom root')
            for item in root:
                if item.tag!=f'{{{CUSTOM}}}property' or len(item)!=1: raise ValueError('invalid custom property')
                value=item[0]; properties.append({'name':item.get('name'),'pid':item.get('pid'),'fmtid':item.get('fmtid'),'type':value.tag.split('}')[-1],'value':value.text or ''})
        return {'part_present':bool(parts),'custom_relationships':len(relationships),'properties':properties}

def main():
    base=ROOT/'.tmp/fresh-properties-libreoffice';base.mkdir(parents=True,exist_ok=True);run=Path(tempfile.mkdtemp(prefix='run-',dir=base))
    js="""import {writeFile} from 'node:fs/promises';import {createPresentation} from './packages/core/dist/index.js';import {generatePptx} from './packages/pptx-exporter/dist/index.js';
const properties=[{name:'Empty',type:'string',value:''},{name:'Flag',type:'boolean',value:false},{name:'Count',type:'integer',value:-2147483648},{name:'Real',type:'real',value:1.25},{name:'Date',type:'date',value:'2020-01-02T03:04:05.000Z'},{name:'Decimal',type:'decimal',value:'9007199254740993.0001'},{name:'Control',type:'string',value:'before'}];
const doc=createPresentation({metadata:{customProperties:properties}});doc.addSlide({id:'fresh-properties'});await writeFile(process.argv[1],(await generatePptx(doc)).bytes);
const updated=[{name:'Flag',type:'boolean',value:true},{name:'Count',type:'integer',value:2147483647},{name:'Real',type:'real',value:0.125},{name:'Date',type:'date',value:'2021-02-03T04:05:06.000Z'},{name:'Decimal',type:'decimal',value:'9007199254740994.0001'},{name:'Control',type:'string',value:'after'}];doc.setMetadata({...doc.metadata,customProperties:updated});await writeFile(process.argv[2],(await generatePptx(doc)).bytes);
doc.setMetadata({...doc.metadata,customProperties:[]});await writeFile(process.argv[3],(await generatePptx(doc)).bytes);const meta={...doc.metadata};delete meta.customProperties;doc.setMetadata(meta);await writeFile(process.argv[4],(await generatePptx(doc)).bytes);"""
    files={name:run/f'{name}.pptx' for name in ('control','edited','empty','absent')}
    subprocess.run(['node','--input-type=module','-e',js,*map(str,files.values())],cwd=ROOT,check=True,timeout=60)
    records=[];logs=[]
    for kind in ('control','edited'):
        source=files[kind];records.append({'kind':kind,'round':0,'path':str(source.relative_to(ROOT)),'sha256':digest(source),**inspect(source)})
        for n in range(1,4):
            out=run/f'{kind}-{n}';out.mkdir();profile=(run/f'profile-{kind}-{n}').as_uri();result=subprocess.run(['soffice',f'-env:UserInstallation={profile}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(out),str(source)],capture_output=True,text=True,check=True,timeout=90);logs.append({'kind':kind,'round':n,'exit_code':result.returncode,'stdout':result.stdout,'stderr':result.stderr});source=out/source.name;records.append({'kind':kind,'round':n,'path':str(source.relative_to(ROOT)),'sha256':digest(source),**inspect(source)})
    auxiliaries=[{'kind':kind,'path':str(files[kind].relative_to(ROOT)),'sha256':digest(files[kind]),**inspect(files[kind])} for kind in ('empty','absent')]
    print(json.dumps({'scope':'Fresh PPTKit public authoring and modification, absent/empty part controls and paired LibreOfficeDev headless resaves; no UI or native authoring.','producer':subprocess.run(['soffice','--version'],capture_output=True,text=True,check=True).stdout.strip(),'records':records,'auxiliaries':auxiliaries,'logs':logs},ensure_ascii=False,indent=2))
if __name__=='__main__': main()
