#!/usr/bin/env python3
"""Observe LPSTR strings through paired LibreOfficeDev headless resaves."""
import hashlib
import importlib.util
import json
import platform
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
BASE = ROOT / '.tmp/custom-lpstr-libreoffice'
ORACLE = ROOT / 'scripts/research/custom-property-create-libreoffice.py'
CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
GUID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'
SAMPLES = [('Unicode','中文 & text'),('Empty',''),('Literal','_x005F_x0041_'),('Emoji','😀')]
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
spec=importlib.util.spec_from_file_location('oracle',ORACLE); oracle=importlib.util.module_from_spec(spec); spec.loader.exec_module(oracle)
def main():
    prior=json.loads((ROOT/'docs/testing/custom-property-create-libreoffice-evidence.json').read_text()); record=next(x for x in prior['records'] if x['round']==0 and x['kind']=='edited'); seed=ROOT/record['path']
    if digest(seed)!=record['sha256']: raise ValueError('seed hash mismatch')
    BASE.mkdir(parents=True,exist_ok=True); run=Path(tempfile.mkdtemp(prefix='run-',dir=BASE)); inputs=run/'input'; inputs.mkdir(); control=inputs/'control.pptx'; edited=inputs/'edited.pptx'
    ET.register_namespace('',CUSTOM); ET.register_namespace('vt',VT); root=ET.Element(f'{{{CUSTOM}}}Properties')
    for index,(name,value) in enumerate(SAMPLES):
        item=ET.SubElement(root,f'{{{CUSTOM}}}property',{'fmtid':GUID,'pid':str(index+2),'name':name}); ET.SubElement(item,f'{{{VT}}}lpstr').text=value
    with zipfile.ZipFile(seed) as before,zipfile.ZipFile(control,'w') as after:
        if 'pptkit-custom-properties-1.xml' not in before.namelist(): raise ValueError('missing seed part')
        for info in before.infolist(): after.writestr(info,ET.tostring(root,encoding='utf-8',xml_declaration=True,short_empty_elements=False) if info.filename=='pptkit-custom-properties-1.xml' else before.read(info.filename))
    replacements=[('中文 &amp; text','Edited 中文 &amp; text'),('','now nonempty'),('_x005F_x0041_','_x005F_x0042_'),('😀','changed 😀')]
    with zipfile.ZipFile(control) as before,zipfile.ZipFile(edited,'w') as after:
        for info in before.infolist():
            data=before.read(info.filename)
            if info.filename=='pptkit-custom-properties-1.xml':
                for left,right in replacements:
                    marker=f'<vt:lpstr>{left}</vt:lpstr>'.encode(); value=f'<vt:lpstr>{right}</vt:lpstr>'.encode()
                    if data.count(marker)!=1: raise ValueError('ambiguous source marker')
                    data=data.replace(marker,value,1)
            after.writestr(info,data)
    with zipfile.ZipFile(control) as a,zipfile.ZipFile(edited) as b:
        if set(a.namelist())!=set(b.namelist()): raise ValueError('entry sets differ')
        changed=sorted(n for n in a.namelist() if a.read(n)!=b.read(n))
    if changed!=['pptkit-custom-properties-1.xml']: raise ValueError('unexpected source changes')
    records=[]; logs=[]
    for n in range(4):
        for kind in ('control','edited'):
            parent=inputs if n==0 else run/f'round-{n}'; path=parent/f'{kind}.pptx'
            if n:
                parent.mkdir(exist_ok=True); previous=(inputs if n==1 else run/f'round-{n-1}')/f'{kind}.pptx'; profile=(run/f'profile-{n}-{kind}').as_uri()
                result=subprocess.run([str(SOFFICE),f'-env:UserInstallation={profile}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(parent),str(previous)],capture_output=True,text=True)
                logs.append({'round':n,'kind':kind,'exitCode':result.returncode,'stdout':result.stdout,'stderr':result.stderr})
                if result.returncode or not path.is_file(): raise ValueError(f'conversion failed: {result.stderr}')
            records.append({'round':n,'kind':kind,'path':str(path.relative_to(ROOT)),'sha256':digest(path),'properties':oracle.inventory(path)})
    manifest={'scope':'Synthetic LPSTR source based on PPTKit-created package, four synthetic value edits and unchanged control, three LibreOfficeDev headless resaves per branch; no UI or PowerPoint.','application':subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(),'platform':platform.platform(),'runnerSha256':digest(Path(__file__)),'oracleSha256':digest(ORACLE),'seedSha256':digest(seed),'sourceChangedParts':changed,'records':records,'logs':logs}
    (ROOT/'docs/testing/custom-lpstr-libreoffice-evidence.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'status':'complete','records':records},ensure_ascii=False))
if __name__=='__main__': main()
