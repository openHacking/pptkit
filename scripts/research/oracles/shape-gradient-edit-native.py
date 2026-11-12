#!/usr/bin/env python3
"""Check one PowerPoint shape edit; artifact hashes alone do not prove provenance."""
import hashlib,json,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/'.tmp/shape-gradient-native-01';N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def inspect(path):
 out=[]
 with zipfile.ZipFile(path) as z:
  for i in range(1,4):
   r=ET.fromstring(z.read(f'ppt/slides/slide{i}.xml'));shapes=r.findall('.//p:sp',N);assert len(shapes)==1
   g=shapes[0].find('p:spPr/a:gradFill',N);lin=g.find('a:lin',N);stops=[]
   for stop in g.findall('a:gsLst/a:gs',N):
    color=stop.find('a:srgbClr',N);alpha=color.find('a:alpha',N);stops.append([stop.get('pos'),color.get('val'),alpha.get('val') if alpha is not None else '100000'])
   xfrm=shapes[0].find('p:spPr/a:xfrm',N);group=r.find('.//p:grpSp/p:grpSpPr/a:xfrm',N)
   out.append({'angle':lin.get('ang'),'scaled':lin.get('scaled'),'stops':stops,'shapeRotation':xfrm.get('rot','0'),'groupRotation':group.get('rot','0') if group is not None else None})
 return out
before=inspect(BASE/'shape-gradient-saved.pptx');after=inspect(BASE/'shape-gradient-edited.pptx')
assert before[0]['angle']=='2700000' and before[0]['scaled']=='1' and before[0]['stops'][0][2]=='70000'
assert after[0]['angle']=='7200000' and after[0]['scaled']=='0' and after[0]['stops'][0][2]=='50000'
assert before[1:]==after[1:] and before[0]['stops'][1]==after[0]['stops'][1]
ax=(BASE/'edited-reopen-ax.txt').read_text();assert '120°' in ax and '50%' in ax and '线性' in ax
out={'scope':'PowerPoint 16.104 edits first shape angle and first stop transparency, Save As, close and reopen once; untouched page gradients and rotations compared. Screenshot and AX hashes alone do not prove provenance. No multi-cycle or general shape fidelity claim.','beforeSha256':sha(BASE/'shape-gradient-saved.pptx'),'afterSha256':sha(BASE/'shape-gradient-edited.pptx'),'uiSha256':sha(BASE/'edited-reopen-ui.png'),'axSha256':sha(BASE/'edited-reopen-ax.txt'),'oracleSha256':sha(Path(__file__)),'before':before,'after':after}
(ROOT/'docs/testing/shape-gradient-native-edit-evidence.json').write_text(json.dumps(out,indent=2)+'\n');print(json.dumps({'before':before[0],'after':after[0]}))
