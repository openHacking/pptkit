#!/usr/bin/env python3
"""Observe PPTKit source-preserving custom edits after headless resaves."""
import hashlib
import importlib.util
import json
import platform
import subprocess
import tempfile
import zipfile
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
BASE = ROOT / '.tmp/custom-property-save-libreoffice'
ORACLE = ROOT / 'scripts/research/custom-properties-libreoffice.py'
def digest(p): return hashlib.sha256(p.read_bytes()).hexdigest()
spec = importlib.util.spec_from_file_location('oracle', ORACLE)
oracle = importlib.util.module_from_spec(spec); spec.loader.exec_module(oracle)
def main():
    prior = json.loads((ROOT / 'docs/testing/authored-custom-properties-libreoffice-evidence.json').read_text())
    source = ROOT / prior['records'][0]['path']
    if digest(source) != prior['records'][0]['sha256']: raise ValueError('source hash mismatch')
    BASE.mkdir(parents=True, exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix='run-', dir=BASE)); input_dir = run / 'input'; input_dir.mkdir()
    control = input_dir / 'control.pptx'; control.write_bytes(source.read_bytes())
    edited = input_dir / 'edited.pptx'
    js = """import { readFile,writeFile } from 'node:fs/promises';
import { importPptx } from './packages/pptx-parser/dist/index.js';
const s=await importPptx(await readFile(process.argv[1]),{projection:'content',customPropertySaveProfile:'scalar-v1'});
try { const values=s.document.metadata.customProperties.map(p=>({...p,value:p.name==='Text'?'Edited 中文 & text':p.name==='Boolean'?true:p.name==='Integer'?42:p.name==='Real'?3.5:'2022-03-04T05:06:07.000Z'}));
s.document.setMetadata({...s.document.metadata,customProperties:values}); const plan=await s.prepareSave({mode:'edit'}); if(plan.status!=='ready')throw Error(JSON.stringify(plan.report));await writeFile(process.argv[2],(await s.save(plan.plan)).bytes); } finally {s.dispose()}"""
    subprocess.run(['node','--input-type=module','-e',js,str(control),str(edited)],cwd=ROOT,check=True)
    with zipfile.ZipFile(control) as a, zipfile.ZipFile(edited) as b:
        if set(a.namelist()) != set(b.namelist()): raise ValueError('entry sets differ')
        changed = sorted(n for n in a.namelist() if a.read(n) != b.read(n))
    if changed != ['docProps/custom.xml']: raise ValueError(f'unexpected source changes: {changed}')
    records=[]; logs=[]
    for round in range(4):
        for kind in ('control','edited'):
            path = (input_dir if round == 0 else run / f'round-{round}') / f'{kind}.pptx'
            if round:
                path.parent.mkdir(exist_ok=True)
                previous = (input_dir if round == 1 else run / f'round-{round-1}') / f'{kind}.pptx'
                profile = (run / f'profile-{round}-{kind}').as_uri()
                process = subprocess.run([str(SOFFICE),f'-env:UserInstallation={profile}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(path.parent),str(previous)],capture_output=True,text=True)
                logs.append({'round':round,'kind':kind,'exitCode':process.returncode,'stdout':process.stdout,'stderr':process.stderr})
                if process.returncode or not path.is_file(): raise ValueError(f'conversion failed: {process.stderr}')
            records.append({'round':round,'kind':kind,'path':str(path.relative_to(ROOT)),'sha256':digest(path),'properties':oracle.inventory(path)})
    result={'scope':'One source-preserving five-value edit and unchanged control; three headless resaves each; not PowerPoint or UI validation.','application':subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(),'platform':platform.platform(),'runnerSha256':digest(Path(__file__)),'oracleSha256':digest(ORACLE),'sourceChangedParts':changed,'records':records,'logs':logs}
    (ROOT / 'docs/testing/custom-property-save-libreoffice-evidence.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'status':'complete','records':records},ensure_ascii=False))
if __name__ == '__main__': main()
