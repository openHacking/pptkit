#!/usr/bin/env python3
"""Observe three-property creation in LibreOfficeDev with an unchanged control."""
import hashlib, json, platform, subprocess, tempfile, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
BASE = ROOT / '.tmp/extended-create-libreoffice'
REL = {'core':'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'extended':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'custom':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'}
DC = 'http://purl.org/dc/elements/1.1/'
EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def inventory(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or len(z.namelist()) != len(set(z.namelist())): raise ValueError('invalid ZIP')
        rels = ET.fromstring(z.read('_rels/.rels')); result = {}
        for kind, relation in REL.items():
            found = [r for r in rels if r.get('Type') == relation]
            if len(found) > 1: raise ValueError('ambiguous relationship')
            if not found: result[kind] = None; continue
            target = found[0].get('Target','').lstrip('/')
            if not target or '..' in target.split('/') or target not in z.namelist(): raise ValueError('invalid relationship target')
            root = ET.fromstring(z.read(target))
            if kind == 'custom':
                if any(len(item) != 1 for item in root): raise ValueError('unsupported custom structure')
                result[kind] = [{'name':item.get('name'), 'type':item[0].tag.split('}')[-1], 'value':''.join(item[0].itertext())} for item in root]
            else:
                tag, ns = ('title', DC) if kind == 'core' else ('Company', EP)
                found = root.findall(f'{{{ns}}}{tag}')
                if len(found)>1: raise ValueError('duplicate property')
                result[kind] = ''.join(found[0].itertext()) if found else None
        return result

def main():
    seed_record = json.loads((ROOT/'docs/testing/core-metadata-libreoffice-evidence.json').read_text())['generated']; seed=ROOT/seed_record['path']
    if digest(seed)!=seed_record['sha256']: raise ValueError('seed hash mismatch')
    BASE.mkdir(parents=True,exist_ok=True); run=Path(tempfile.mkdtemp(prefix='run-',dir=BASE)); inputs=run/'input'; inputs.mkdir(); control=inputs/'control.pptx'; edited=inputs/'edited.pptx'
    with zipfile.ZipFile(seed) as src, zipfile.ZipFile(control,'w') as dst:
        for info in src.infolist():
            if info.filename in ('docProps/core.xml','docProps/app.xml'): continue
            payload=src.read(info.filename)
            if info.filename=='_rels/.rels':
                for marker in [b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',b'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>']:
                    if payload.count(marker)!=1: raise ValueError('unexpected seed relationship')
                    payload=payload.replace(marker,b'')
            elif info.filename=='[Content_Types].xml':
                for marker in [b'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',b'<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>']:
                    if payload.count(marker)!=1: raise ValueError('unexpected seed declaration')
                    payload=payload.replace(marker,b'')
            dst.writestr(info,payload)
    if inventory(control)!={'core':None,'extended':None,'custom':None}: raise ValueError('unexpected control metadata')
    js="""import {readFile,writeFile} from 'node:fs/promises';import {importPptx} from './packages/pptx-parser/dist/index.js';
const s=await importPptx(await readFile(process.argv[1]),{projection:'content',metadataSaveProfile:'extended-create-v12',customPropertySaveProfile:'create-and-append-v5'});
try{s.document.setMetadata({...s.document.metadata,title:'Created 中文',company:'Company & 中文',customProperties:[{name:'Count',type:'integer',value:42}]});const p=await s.prepareSave({mode:'edit',customPropertyAdditions:[0]});if(p.status!=='ready')throw Error(JSON.stringify(p.report));await writeFile(process.argv[2],(await s.save(p.plan)).bytes)}finally{s.dispose()}"""
    subprocess.run(['node','--input-type=module','-e',js,str(control),str(edited)],cwd=ROOT,check=True)
    with zipfile.ZipFile(control) as a,zipfile.ZipFile(edited) as b:
        added=sorted(set(b.namelist())-set(a.namelist())); removed=sorted(set(a.namelist())-set(b.namelist())); changed=sorted(n for n in a.namelist() if a.read(n)!=b.read(n))
    if added!=['pptkit-core-properties-1.xml','pptkit-custom-properties-1.xml','pptkit-extended-properties-1.xml'] or removed or changed!=['[Content_Types].xml','_rels/.rels']: raise ValueError('unexpected package delta')
    records=[]; logs=[]
    for n in range(4):
        for kind in ('control','edited'):
            parent=inputs if n==0 else run/f'round-{n}'; path=parent/f'{kind}.pptx'
            if n:
                parent.mkdir(exist_ok=True); previous=(inputs if n==1 else run/f'round-{n-1}')/f'{kind}.pptx'
                if not previous.is_file(): logs.append({'round':n,'kind':kind,'status':'skipped'}); continue
                result=subprocess.run([str(SOFFICE),f'-env:UserInstallation={(run/f"profile-{n}-{kind}").as_uri()}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(parent),str(previous)],capture_output=True,text=True)
                logs.append({'round':n,'kind':kind,'status':'saved' if result.returncode==0 and path.is_file() else 'failed','exitCode':result.returncode,'stdout':result.stdout,'stderr':result.stderr})
                if result.returncode or not path.is_file(): continue
            records.append({'round':n,'kind':kind,'path':str(path.relative_to(ROOT)),'sha256':digest(path),'metadata':inventory(path)})
    evidence={'scope':'One synthetic missing core/extended/custom control and PPTKit combined creation; LibreOfficeDev headless resaves, no interactive UI or PowerPoint.','application':subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(),'platform':platform.platform(),'runnerSha256':digest(Path(__file__)),'seed':seed_record,'sourceAddedParts':added,'sourceChangedParts':changed,'records':records,'logs':logs}
    (ROOT/'docs/testing/extended-create-libreoffice-evidence.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'records':records,'logs':logs},ensure_ascii=False))
if __name__=='__main__': main()
