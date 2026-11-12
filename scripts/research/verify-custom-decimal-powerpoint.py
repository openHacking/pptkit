#!/usr/bin/env python3
"""Verify retained files from one PowerPoint decimal-property comparison."""
import hashlib,json,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]; BASE=ROOT/'.tmp/custom-decimal-powerpoint'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
CUSTOM='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
VT='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
files=['control.pptx','input.pptx','control-saved.pptx','saved1.pptx'];records=[]
for name in files:
 path=BASE/name
 with zipfile.ZipFile(path) as z:
  if z.testzip() is not None or len(z.namelist())!=len(set(z.namelist())): raise ValueError('invalid ZIP')
  relations=[r for r in ET.fromstring(z.read('_rels/.rels')) if r.get('Type')==REL]
  props=[]
  if relations:
   if len(relations)!=1 or relations[0].get('TargetMode','Internal')!='Internal': raise ValueError('ambiguous custom properties')
   part=relations[0].get('Target','').lstrip('/')
   if not part or '..' in part.split('/'): raise ValueError('unsafe target')
   root=ET.fromstring(z.read(part))
   if root.tag!=f'{{{CUSTOM}}}Properties': raise ValueError('wrong custom root')
   for child in root:
    if len(child)!=1 or child[0].tag!=f'{{{VT}}}decimal': raise ValueError('unexpected property')
    props.append([child.get('name'),child.get('pid'),child[0].text])
  if not relations and any('custom-properties' in n.lower() or n.lower() == 'docprops/custom.xml' for n in z.namelist()): raise ValueError('orphan custom part')
  types=ET.fromstring(z.read('[Content_Types].xml'))
  if not relations and any(item.get('ContentType') == 'application/vnd.openxmlformats-officedocument.custom-properties+xml' for item in types): raise ValueError('custom content type retained')
  slide=ET.fromstring(z.read('ppt/slides/slide1.xml'))
  if ''.join(x.text or '' for x in slide.iter(f'{{{A}}}t'))!='Metadata control slide': raise ValueError('slide text changed')
  if name=='control.pptx': expected=[['Normal','2','1.25'],['Fraction','3','0.125'],['Inexact','4','0.1'],['Large','5','9007199254740992']]
  elif name=='input.pptx': expected=[['Normal','2','2.5'],['Fraction','3','0.5'],['Inexact','4','0.1'],['Large','5','4503599627370496']]
  else: expected=[]
  if props!=expected or bool(relations)!=(name in ('control.pptx','input.pptx')): raise ValueError('unexpected properties')
  records.append({'file':str(path.relative_to(ROOT)),'sha256':sha(path),'customRelationshipCount':len(relations),'properties':props})
with zipfile.ZipFile(BASE/'control.pptx') as a,zipfile.ZipFile(BASE/'input.pptx') as b:
 if set(a.namelist())!=set(b.namelist()) or sorted(n for n in a.namelist() if a.read(n)!=b.read(n))!=['pptkit-custom-properties-1.xml']: raise ValueError('unexpected source changes')
output={'scope':'One synthetic decimal sample and unchanged control, each saved once in PowerPoint. Edited save reopened with slide text visible; UI observations are not independently attested by ZIP hashes.','application':'PowerPoint 16.104 (25121423)','verifierSha256':sha(Path(__file__)),'records':records}
destination=ROOT/'docs/testing/custom-decimal-powerpoint-evidence.json'
if destination.exists():
 if json.loads(destination.read_text())!=output: raise ValueError('retained evidence mismatch')
else: destination.write_text(json.dumps(output,ensure_ascii=False,indent=2)+'\n')
print('Verified four ZIPs, slide text, source isolation and loss of both custom-property relationships.')
