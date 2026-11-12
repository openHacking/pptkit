#!/usr/bin/env python3
"""Inspect a narrow first-open PowerPoint slideshow and SVG gradient comparison."""
import hashlib,json,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
from PIL import Image
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/'.tmp/shape-gradient-native-01'
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def digest(p):return hashlib.sha256(p.read_bytes()).hexdigest()
for name in ('shape-gradient.pptx','shape-gradient-saved.pptx'):
 with zipfile.ZipFile(BASE/name) as z:
  for i in range(1,4):
   root=ET.fromstring(z.read(f'ppt/slides/slide{i}.xml')); shapes=root.findall('.//p:sp',N);assert len(shapes)==1
   grad=shapes[0].find('p:spPr/a:gradFill',N);lin=grad.find('a:lin',N);assert lin.get('ang')=='2700000' and lin.get('scaled')=='1'
   stops=grad.findall('a:gsLst/a:gs',N);alpha=stops[0].find('a:srgbClr/a:alpha',N);assert alpha.get('val')==('35000' if i==3 else '70000')
   xfrm=shapes[0].find('p:spPr/a:xfrm',N);assert (xfrm.get('rot') or '0')==('1800000' if i==2 else '0')
   if i==3:assert root.find('.//p:grpSp/p:grpSpPr/a:xfrm',N).get('rot')=='1800000'
points=[(480,270),(300,180),(660,360),(300,270),(660,270),(480,180),(480,360),(100,100)]
observed={}
for i in range(1,4):
 a=Image.open(BASE/f'native-{i}.png').convert('RGB');b=Image.open(BASE/f'svg-{i}.png').convert('RGB');assert a.size==(1920,1080) and b.size==(960,540)
 observed[str(i)]={'native':[a.getpixel((2*x,2*y)) for x,y in points],'svg':[b.getpixel((x,y)) for x,y in points]}
 reopen=Image.open(BASE/f'reopen-{i}.png').convert('RGB');assert reopen.size==a.size and [reopen.getpixel((2*x,2*y)) for x,y in points]==observed[str(i)]['native']
 for v in observed[str(i)].values():assert v[1][0]>v[2][0] and v[1][2]<v[2][2] and v[7]==(255,255,255)
for kind in ('native','svg'):
 assert observed['1'][kind][3][0]>observed['1'][kind][4][0]+50
 assert observed['2'][kind][3][0]<observed['2'][kind][4][0]
 assert observed['3'][kind][0][1]>observed['2'][kind][0][1]+70
assert max(abs(x-y) for x,y in zip(observed['1']['native'][0],observed['1']['svg'][0]))>=30
out={'scope':'PowerPoint 16.104 first-open slideshow: one scaled 45-degree two-stop rectangle, a 30-degree rotated rectangle, and a rotated half-opacity group. Sampled colors and directions are evidence only for this fixture; a Save As/close/reopen preserves selected XML and pixel samples. Screenshot hashes alone do not prove provenance. No native gradient edit.', 'pptxSha256':digest(BASE/'shape-gradient.pptx'),'savedPptxSha256':digest(BASE/'shape-gradient-saved.pptx'), 'reopenScreenshots':[digest(BASE/f'reopen-{i}.png') for i in range(1,4)],'producerSha256':digest(ROOT/'scripts/prepare-shape-gradient-native.mjs'),'oracleSha256':digest(Path(__file__)),'nativeScreenshots':[digest(BASE/f'native-{i}.png') for i in range(1,4)],'svgScreenshots':[digest(BASE/f'svg-{i}.png') for i in range(1,4)],'samplePoints':points,'observed':observed}
(ROOT/'docs/testing/shape-gradient-native-evidence.json').write_text(json.dumps(out,indent=2)+'\n');print(json.dumps(observed,indent=2))
