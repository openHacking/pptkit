#!/usr/bin/env python3
"""Inspect one PowerPoint edit/save/reopen sample; it does not attest UI provenance."""
import hashlib,json,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/'.tmp/background-gradient-native-01'
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def hashfile(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def get(z,i):
 r=ET.fromstring(z.read(f'ppt/slides/slide{i}.xml'));g=r.find('p:cSld/p:bg/p:bgPr/a:gradFill',N);assert g is not None
 lin=g.find('a:lin',N);stops=[]
 for s in g.findall('a:gsLst/a:gs',N):
  c=s.find('a:srgbClr',N);assert c is not None; alpha=c.find('a:alpha',N);stops.append([s.get('pos'),c.get('val'),alpha.get('val') if alpha is not None else '100000'])
 return {'angle':lin.get('ang'),'scaled':lin.get('scaled'),'stops':stops}
with zipfile.ZipFile(BASE/'gradient.pptx') as original,zipfile.ZipFile(BASE/'gradient-edited.pptx') as edited:
 before=[get(original,i) for i in range(1,7)];after=[get(edited,i) for i in range(1,7)]
 assert before[3]=={'angle':'2700000','scaled':'1','stops':[['0','FF0000','100000'],['100000','0000FF','100000']]}
 assert after[3]=={'angle':'7200000','scaled':'0','stops':[['0','FF0000','70000'],['100000','0000FF','100000']]}
 assert all(before[i]==after[i] for i in [0,1,2,4,5])
ax=(BASE/'edited-reopen-ax.txt').read_text();assert '120°' in ax and '30%' in ax and '线性' in ax and '第 4 张幻灯片，共 6 张' in ax
result={'scope':'PowerPoint 16.104: edit slide 4 angle and first stop transparency, save as, close and reopen once. Other slide gradient data compared semantically; screenshots and AX text alone do not independently establish provenance. No multiple reopen cycles or general gradient support established.', 'sourceSha256':hashfile(BASE/'gradient.pptx'),'editedSha256':hashfile(BASE/'gradient-edited.pptx'),'reopenScreenshotSha256':hashfile(BASE/'edited-reopen-ui.png'),'reopenAxSha256':hashfile(BASE/'edited-reopen-ax.txt'),'oracleSha256':hashfile(Path(__file__)),'before':before,'after':after}
(ROOT/'docs/testing/background-gradient-native-edit-evidence.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps({'before':before[3],'after':after[3]},indent=2))
