#!/usr/bin/env python3
"""Observe one explicit source keyword edit and an unchanged control in LibreOfficeDev."""
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
BASE = ROOT / '.tmp/core-keywords-libreoffice'
CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
DC = 'http://purl.org/dc/elements/1.1/'
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def inventory(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or len(z.namelist()) != len(set(z.namelist())): raise ValueError('invalid ZIP')
        root = ET.fromstring(z.read('docProps/core.xml'))
        result = {}
        for field, namespace in [('title', DC), ('creator', DC), ('revision', CP), ('keywords', CP)]:
            matches = root.findall(f'{{{namespace}}}{field}')
            if len(matches) > 1: raise ValueError('duplicate core property')
            result[field] = ''.join(matches[0].itertext()) if matches else None
        return result

def main():
    prior = json.loads((ROOT / 'docs/testing/core-metadata-libreoffice-evidence.json').read_text()); original = ROOT / prior['generated']['path']
    if digest(original) != prior['generated']['sha256'] or inventory(original)['keywords'] != 'alpha, beta': raise ValueError('seed hash or keywords mismatch')
    BASE.mkdir(parents=True, exist_ok=True); run = Path(tempfile.mkdtemp(prefix='run-', dir=BASE)); inputs = run / 'input'; inputs.mkdir()
    control = inputs / 'control.pptx'; control.write_bytes(original.read_bytes()); edited = inputs / 'edited.pptx'
    js = """import { readFile,writeFile } from 'node:fs/promises';
import { importPptx } from './packages/pptx-parser/dist/index.js';
const s=await importPptx(await readFile(process.argv[1]),{projection:'content',metadataSaveProfile:'core-keywords-v3'});
try{s.document.setMetadata({...s.document.metadata,keywords:['New, value & 中文']});const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));await writeFile(process.argv[2],(await s.save(p.plan)).bytes);}finally{s.dispose()}"""
    subprocess.run(['node','--input-type=module','-e',js,str(control),str(edited)],cwd=ROOT,check=True)
    with zipfile.ZipFile(control) as a, zipfile.ZipFile(edited) as b:
        if set(a.namelist()) != set(b.namelist()): raise ValueError('entry sets differ')
        changed = sorted(n for n in a.namelist() if a.read(n) != b.read(n))
    if changed != ['docProps/core.xml'] or inventory(edited)['keywords'] != 'New, value & 中文': raise ValueError('unexpected source edit')
    records=[]; logs=[]
    for n in range(4):
        for kind in ('control','edited'):
            parent = inputs if n == 0 else run / f'round-{n}'; path = parent / f'{kind}.pptx'
            if n:
                parent.mkdir(exist_ok=True); previous = (inputs if n == 1 else run / f'round-{n-1}') / f'{kind}.pptx'
                profile=(run / f'profile-{n}-{kind}').as_uri()
                result=subprocess.run([str(SOFFICE),f'-env:UserInstallation={profile}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(parent),str(previous)],capture_output=True,text=True)
                logs.append({'round':n,'kind':kind,'exitCode':result.returncode,'stdout':result.stdout,'stderr':result.stderr})
                if result.returncode or not path.is_file(): raise ValueError(f'conversion failed: {result.stderr}')
            records.append({'round':n,'kind':kind,'path':str(path.relative_to(ROOT)),'sha256':digest(path),'metadata':inventory(path)})
    result={'scope':'One source keyword edit and unchanged control, three LibreOfficeDev headless resaves per branch; no interactive UI or PowerPoint.','application':subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(),'platform':platform.platform(),'runnerSha256':digest(Path(__file__)),'sourceChangedParts':changed,'records':records,'logs':logs}
    (ROOT / 'docs/testing/core-keywords-libreoffice-evidence.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'status':'complete','records':records},ensure_ascii=False))
if __name__ == '__main__': main()
