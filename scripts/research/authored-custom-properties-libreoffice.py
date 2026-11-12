#!/usr/bin/env python3
"""Observe actual PPTKit-authored custom properties through headless resaves."""
import hashlib
import importlib.util
import json
import platform
import subprocess
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
BASE = ROOT / '.tmp/authored-custom-properties-libreoffice'
ORACLE = ROOT / 'scripts/research/custom-properties-libreoffice.py'

def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
spec = importlib.util.spec_from_file_location('custom_oracle', ORACLE)
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)

def main():
    BASE.mkdir(parents=True, exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix='run-', dir=BASE))
    source = run / 'authored.pptx'
    code = """import { writeFile } from 'node:fs/promises';
import { createPresentation } from './packages/core/dist/index.js';
import { generatePptx } from './packages/pptx-exporter/dist/index.js';
const p=createPresentation({metadata:{customProperties:[
{name:'Text',type:'string',value:'中文 & text'}, {name:'Boolean',type:'boolean',value:false},
{name:'Integer',type:'integer',value:-2147483648}, {name:'Real',type:'real',value:1.25},
{name:'Date',type:'date',value:'2020-01-02T03:04:05.000Z'}]}});
p.addSlide({id:'slide'}).addElement({type:'text',content:'Authored custom property control',box:{x:20,y:20,width:400,height:80}});
await writeFile(process.argv[1],(await generatePptx(p)).bytes);"""
    subprocess.run(['node', '--input-type=module', '-e', code, str(source)], cwd=ROOT, check=True)
    records = [{'round': 0, 'path': str(source.relative_to(ROOT)), 'sha256': digest(source), 'properties': oracle.inventory(source)}]
    expected = [('Text','lpwstr','中文 & text'), ('Boolean','bool','false'), ('Integer','i4','-2147483648'), ('Real','r8','1.25'), ('Date','filetime','2020-01-02T03:04:05.000Z')]
    actual = [(p['name'],p['type'].split('}')[-1],p['value']) for p in records[0]['properties']]
    if actual != expected: raise ValueError('authored source inventory mismatch')
    current = source
    logs = []
    for number in range(1,4):
        out = run / f'round-{number}'; out.mkdir()
        profile = (run / f'profile-{number}').as_uri()
        result = subprocess.run([str(SOFFICE), f'-env:UserInstallation={profile}', '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', str(out), str(current)], capture_output=True, text=True)
        logs.append({'round': number, 'exitCode': result.returncode, 'stdout': result.stdout, 'stderr': result.stderr})
        if result.returncode != 0: raise RuntimeError(f'conversion failed: {result.stderr}')
        current = out / current.name
        if not current.is_file(): raise ValueError('missing converted file')
        records.append({'round': number, 'path': str(current.relative_to(ROOT)), 'sha256': digest(current), 'properties': oracle.inventory(current)})
    manifest = {'scope':'PPTKit fresh authoring, one sample, LibreOfficeDev headless three resaves; not a PowerPoint UI test.', 'application': subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(), 'platform':platform.platform(), 'runnerSha256':digest(Path(__file__)), 'oracleSha256':digest(ORACLE), 'records':records, 'logs':logs}
    (ROOT / 'docs/testing/authored-custom-properties-libreoffice-evidence.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'status':'complete','records':records},ensure_ascii=False))

if __name__ == '__main__': main()
