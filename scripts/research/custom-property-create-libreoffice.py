#!/usr/bin/env python3
"""Compare a missing-part source and a PPTKit-created custom part through headless resaves."""
import hashlib
import json
import platform
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
BASE = ROOT / '.tmp/custom-property-create-libreoffice'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'

def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def inventory(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if len(names) != len(set(names)): raise ValueError('duplicate ZIP entries')
        rels = ET.fromstring(z.read('_rels/.rels'))
        matches = [r for r in rels if r.get('Type') == REL]
        if not matches: return []
        if len(matches) != 1 or matches[0].get('TargetMode') == 'External': raise ValueError('ambiguous relationship')
        target = matches[0].get('Target', '').lstrip('/')
        if '..' in target.split('/') or target not in names: raise ValueError('invalid target')
        root = ET.fromstring(z.read(target))
        if root.tag != f'{{{CUSTOM}}}Properties': raise ValueError('unexpected property root')
        return [{'name': p.get('name'), 'pid': p.get('pid'), 'type': p[0].tag.split('}')[-1], 'value': ''.join(p[0].itertext())} for p in root if len(p) == 1]

def main():
    prior = json.loads((ROOT / 'docs/testing/core-metadata-libreoffice-evidence.json').read_text())
    source = ROOT / prior['generated']['path']
    if digest(source) != prior['generated']['sha256'] or inventory(source): raise ValueError('source identity or absence mismatch')
    BASE.mkdir(parents=True, exist_ok=True); run = Path(tempfile.mkdtemp(prefix='run-', dir=BASE)); inputs = run / 'input'; inputs.mkdir()
    control = inputs / 'control.pptx'; control.write_bytes(source.read_bytes()); edited = inputs / 'edited.pptx'
    js = """import { readFile,writeFile } from 'node:fs/promises';
import { importPptx } from './packages/pptx-parser/dist/index.js';
const s=await importPptx(await readFile(process.argv[1]),{projection:'content',customPropertySaveProfile:'create-and-append-v5'});
try {const customProperties=[{name:'Text',type:'string',value:'中文 & text'},{name:'Boolean',type:'boolean',value:true},{name:'Integer',type:'integer',value:42},{name:'Real',type:'real',value:3.5},{name:'Date',type:'date',value:'2020-01-02T03:04:05.000Z'}];
s.document.setMetadata({...s.document.metadata,customProperties});const p=await s.prepareSave({mode:'edit',customPropertyAdditions:[0,1,2,3,4]});if(p.status!=='ready')throw Error(JSON.stringify(p.report));await writeFile(process.argv[2],(await s.save(p.plan)).bytes);}finally{s.dispose()}"""
    subprocess.run(['node','--input-type=module','-e',js,str(control),str(edited)],cwd=ROOT,check=True)
    with zipfile.ZipFile(control) as a, zipfile.ZipFile(edited) as b:
        added = sorted(set(b.namelist()) - set(a.namelist())); removed = sorted(set(a.namelist()) - set(b.namelist()))
        changed = sorted(n for n in a.namelist() if a.read(n) != b.read(n))
    if added != ['pptkit-custom-properties-1.xml'] or removed or changed != ['[Content_Types].xml','_rels/.rels']: raise ValueError('unexpected source changes')
    records=[]; logs=[]
    for round_number in range(4):
        for kind in ('control','edited'):
            parent = inputs if round_number == 0 else run / f'round-{round_number}'
            path = parent / f'{kind}.pptx'
            if round_number:
                parent.mkdir(exist_ok=True); previous = (inputs if round_number == 1 else run / f'round-{round_number-1}') / f'{kind}.pptx'
                profile=(run / f'profile-{round_number}-{kind}').as_uri()
                result=subprocess.run([str(SOFFICE),f'-env:UserInstallation={profile}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(parent),str(previous)],capture_output=True,text=True)
                logs.append({'round':round_number,'kind':kind,'exitCode':result.returncode,'stdout':result.stdout,'stderr':result.stderr})
                if result.returncode or not path.is_file(): raise ValueError(f'conversion failed: {result.stderr}')
            records.append({'round':round_number,'kind':kind,'path':str(path.relative_to(ROOT)),'sha256':digest(path),'properties':inventory(path)})
    manifest={'scope':'One PPTKit-created custom part and unchanged control, three LibreOfficeDev headless resaves each; no PowerPoint or UI interaction.','application':subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(),'platform':platform.platform(),'runnerSha256':digest(Path(__file__)),'sourceChangedParts':changed,'sourceAddedParts':added,'records':records,'logs':logs}
    (ROOT / 'docs/testing/custom-property-create-libreoffice-evidence.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'status':'complete','records':records},ensure_ascii=False))
if __name__ == '__main__': main()
