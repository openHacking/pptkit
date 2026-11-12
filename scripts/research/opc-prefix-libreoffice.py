#!/usr/bin/env python3
"""Isolate OPC prefix and XML declaration variables for LibreOfficeDev."""
import hashlib,json,platform,subprocess,tempfile,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];BASE=ROOT/'.tmp/opc-prefix-libreoffice';SOFFICE=Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
REL='http://schemas.openxmlformats.org/package/2006/relationships';CT='http://schemas.openxmlformats.org/package/2006/content-types'
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def tree(node):return (node.tag,tuple(sorted(node.attrib.items())),node.text,tuple(tree(x) for x in node))
def prefixed(payload,kind):
 text=payload.decode('utf-8');ns=REL if kind=='rel' else CT;prefix='r' if kind=='rel' else 't';names=['Relationships','Relationship'] if kind=='rel' else ['Types','Default','Override']
 target=f'xmlns="{ns}"'
 if text.count(target)!=1:raise ValueError('unexpected namespace declaration')
 text=text.replace(target,f'xmlns:{prefix}="{ns}"',1)
 for name in names:text=text.replace(f'<{name}',f'<{prefix}:{name}').replace(f'</{name}>',f'</{prefix}:{name}>')
 return text.encode('utf-8')
def make(source,dest,variant):
 with zipfile.ZipFile(source) as a,zipfile.ZipFile(dest,'w') as b:
  for info in a.infolist():
   payload=a.read(info.filename)
   if info.filename=='_rels/.rels' and variant in ('rel-prefix','both-prefix'):payload=prefixed(payload,'rel')
   if info.filename=='[Content_Types].xml' and variant in ('types-prefix','both-prefix'):payload=prefixed(payload,'types')
   if info.filename in ('_rels/.rels','[Content_Types].xml') and variant=='declaration':
    marker=b' standalone="yes"';
    if payload.count(marker)!=1:raise ValueError('unexpected XML declaration')
    payload=payload.replace(marker,b'',1)
   b.writestr(info,payload)
def main():
 prior=json.loads((ROOT/'docs/testing/combined-property-create-libreoffice-evidence.json').read_text());inputs={r['kind']:r for r in prior['records'] if r['round']==0};BASE.mkdir(parents=True,exist_ok=True);run=Path(tempfile.mkdtemp(prefix='run-',dir=BASE));records=[]
 for kind,record in inputs.items():
  source=ROOT/record['path'];
  if sha(source)!=record['sha256']:raise ValueError('seed hash mismatch')
  for variant in ('baseline','rel-prefix','types-prefix','both-prefix','declaration'):
   folder=run/f'{kind}-{variant}';folder.mkdir();file=folder/'input.pptx';make(source,file,variant)
   with zipfile.ZipFile(source) as a,zipfile.ZipFile(file) as b:
    if set(a.namelist())!=set(b.namelist()):raise ValueError('entry set changed')
    changed=sorted(n for n in a.namelist() if a.read(n)!=b.read(n))
    if any(n not in ('_rels/.rels','[Content_Types].xml') for n in changed):raise ValueError('unapproved change')
    for n in changed:
     if tree(ET.fromstring(a.read(n)))!=tree(ET.fromstring(b.read(n))):raise ValueError('declaration semantics differ')
   output=folder/'output';output.mkdir();r=subprocess.run([str(SOFFICE),f'-env:UserInstallation={(folder/"profile").as_uri()}','--headless','--convert-to','pptx:Impress Office Open XML','--outdir',str(output),str(file)],capture_output=True,text=True)
   result=output/'input.pptx';records.append({'kind':kind,'variant':variant,'path':str(file.relative_to(ROOT)),'sha256':sha(file),'changedParts':changed,'exitCode':r.returncode,'outputExists':result.is_file(),'outputSha256':sha(result) if result.is_file() else None,'stdout':r.stdout,'stderr':r.stderr})
 evidence={'scope':'Two synthetic source branches; independently vary OPC root/child namespace prefixes and XML declaration standalone, one LibreOfficeDev headless conversion per input, no UI or PowerPoint.','application':subprocess.run([str(SOFFICE),'--version'],capture_output=True,text=True,check=True).stdout.strip(),'platform':platform.platform(),'runnerSha256':sha(Path(__file__)),'seedEvidenceSha256':sha(ROOT/'docs/testing/combined-property-create-libreoffice-evidence.json'),'records':records}
 (ROOT/'docs/testing/opc-prefix-libreoffice-evidence.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2)+'\n');print(json.dumps([{'kind':r['kind'],'variant':r['variant'],'exitCode':r['exitCode'],'outputExists':r['outputExists']} for r in records]))
if __name__=='__main__':main()
