#!/usr/bin/env python3
"""Check retained PowerPoint source/save artifacts, independent of PPTKit parsing."""
import hashlib, json, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]
BASE=ROOT/'.tmp/core-insert-powerpoint'
DC='http://purl.org/dc/elements/1.1/'
CP='http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
DCT='http://purl.org/dc/terms/'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
files=['round1.pptx','saved1.pptx','saved2.pptx','saved3.pptx']
expected={'title':'PPTKit Presentation','subject':'','description':'Inserted & 中文','language':'fr-FR','creator':'Sentinel Author','keywords':'alpha, beta','revision':'7'}
records=[]
for name in files:
 p=BASE/name
 with zipfile.ZipFile(p) as z:
  if z.testzip() is not None or len(z.namelist())!=len(set(z.namelist())): raise ValueError('invalid ZIP')
  root=ET.fromstring(z.read('docProps/core.xml')); actual={}
  for field in expected:
   ns=CP if field in ('keywords','revision') else DC
   elements=root.findall(f'{{{ns}}}{field}')
   if len(elements)!=1: raise ValueError(f'{name}: missing or duplicate {field}')
   actual[field]=''.join(elements[0].itertext())
  if actual!=expected: raise ValueError(f'{name}: metadata changed: {actual}')
  slide=ET.fromstring(z.read('ppt/slides/slide1.xml'))
  if ''.join(x.text or '' for x in slide.iter(f'{{{A}}}t'))!='Metadata control slide': raise ValueError(f'{name}: slide changed')
  records.append({'file':str(p.relative_to(ROOT)),'sha256':sha(p),'metadata':actual,'lastModifiedByPresent':root.find(f'{{{CP}}}lastModifiedBy') is not None,'modifiedPresent':root.find(f'{{{DCT}}}modified') is not None})
if records[0]['lastModifiedByPresent'] or records[0]['modifiedPresent'] or not all(r['lastModifiedByPresent'] and r['modifiedPresent'] for r in records[1:]): raise ValueError('unexpected author/time properties')
with zipfile.ZipFile(BASE/files[0]) as a, zipfile.ZipFile(ROOT/'.tmp/core-insert-libreoffice/run-j_hbqn0l/input/control.pptx') as b:
 if set(a.namelist())!=set(b.namelist()) or sorted(n for n in a.namelist() if a.read(n)!=b.read(n))!=['docProps/core.xml']: raise ValueError('unexpected source changes')
output={'scope':'One inserted core-metadata sample, PowerPoint three save-as and close/reopen observations; no metadata UI or editing. UI observations are not independently attested by ZIP hashes.','application':'Microsoft PowerPoint 16.104 (25121423)','oracleSha256':sha(Path(__file__)),'records':records}
path=ROOT/'docs/testing/core-insert-powerpoint-evidence.json'
if path.exists():
 old=json.loads(path.read_text())
 if old!=output: raise ValueError('retained artifacts or verifier differ from evidence')
else: path.write_text(json.dumps(output,ensure_ascii=False,indent=2)+'\n')
print('Verified four ZIP inventories and source isolation.')
