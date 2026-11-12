#!/usr/bin/env python3
"""Compare a missing-field control and source insertion under headless LibreOffice."""
import hashlib, json, platform, subprocess, tempfile, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
DC = 'http://purl.org/dc/elements/1.1/'
CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
FIELDS = [('title', DC), ('creator', DC), ('subject', DC), ('description', DC), ('language', DC), ('revision', CP), ('keywords', CP)]
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def inventory(path):
    with zipfile.ZipFile(path) as archive:
        if archive.testzip() or len(archive.namelist()) != len(set(archive.namelist())): raise ValueError('invalid ZIP')
        root = ET.fromstring(archive.read('docProps/core.xml')); result = {}
        for name, ns in FIELDS:
            matches = root.findall(f'{{{ns}}}{name}')
            if len(matches) > 1: raise ValueError('duplicate core property')
            result[name] = ''.join(matches[0].itertext()) if matches else None
        return result

def main():
    seed = json.loads((ROOT / 'docs/testing/core-metadata-libreoffice-evidence.json').read_text())['generated']
    original = ROOT / seed['path']
    if digest(original) != seed['sha256']: raise ValueError('seed hash mismatch')
    base = ROOT / '.tmp/core-insert-libreoffice'; base.mkdir(parents=True, exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix='run-', dir=base)); inputs = run / 'input'; inputs.mkdir()
    control = inputs / 'control.pptx'; edited = inputs / 'edited.pptx'
    with zipfile.ZipFile(original) as source, zipfile.ZipFile(control, 'w') as output:
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename == 'docProps/core.xml':
                root = ET.fromstring(payload)
                for name in ('title','subject','description','language'):
                    for element in root.findall(f'{{{DC}}}{name}'): root.remove(element)
                payload = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            output.writestr(info,payload)
    js = """import {readFile,writeFile} from 'node:fs/promises';
import {importPptx} from './packages/pptx-parser/dist/index.js';
const s=await importPptx(await readFile(process.argv[1]),{projection:'content',metadataSaveProfile:'core-insert-v5'});
try{s.document.setMetadata({...s.document.metadata,subject:'',description:'Inserted & 中文',language:'fr-FR'});const p=await s.prepareSave({mode:'edit',coreMetadataInsertions:['title','subject','description','language']});if(p.status!=='ready')throw Error(JSON.stringify(p.report));await writeFile(process.argv[2],(await s.save(p.plan)).bytes)}finally{s.dispose()}"""
    subprocess.run(['node','--input-type=module','-e',js,str(control),str(edited)],cwd=ROOT,check=True)
    with zipfile.ZipFile(control) as a, zipfile.ZipFile(edited) as b:
        if set(a.namelist()) != set(b.namelist()): raise ValueError('entry sets differ')
        changed=sorted(n for n in a.namelist() if a.read(n) != b.read(n))
    if changed != ['docProps/core.xml']: raise ValueError('unexpected source edit')
    records=[]; logs=[]
    for n in range(4):
        for kind in ('control','edited'):
            parent=inputs if n==0 else run/f'round-{n}'; path=parent/f'{kind}.pptx'
            if n:
                parent.mkdir(exist_ok=True); previous=(inputs if n==1 else run/f'round-{n-1}')/f'{kind}.pptx'
                result=subprocess.run([str(SOFFICE),f'-env:UserInstallation={(run/f"profile-{n}-{kind}").as_uri()}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(parent),str(previous)],capture_output=True,text=True)
                logs.append({'round':n,'kind':kind,'exitCode':result.returncode,'stdout':result.stdout,'stderr':result.stderr})
                if result.returncode or not path.is_file(): raise ValueError(f'conversion failed: {result.stderr}')
            records.append({'round':n,'kind':kind,'path':str(path.relative_to(ROOT)),'sha256':digest(path),'metadata':inventory(path)})
    evidence={'scope':'Missing core fields and inserted values, each with three LibreOfficeDev headless resaves; no UI or PowerPoint.','application':subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(),'platform':platform.platform(),'runnerSha256':digest(Path(__file__)),'sourceChangedParts':changed,'records':records,'logs':logs}
    (ROOT/'docs/testing/core-insert-libreoffice-evidence.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps(records,ensure_ascii=False))
if __name__=='__main__': main()
