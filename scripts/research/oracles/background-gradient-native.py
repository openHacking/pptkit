#!/usr/bin/env python3
"""Check scoped native slideshow observations and the corrected SVG diagonal."""
import hashlib,json,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
from PIL import Image
ROOT=Path(__file__).resolve().parents[3]
NATIVE=ROOT/'.tmp/background-gradient-native-01'
PREVIEW=ROOT/'.tmp/background-gradient-native-02'
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def digest(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def diff(a,b):return max(abs(x-y) for x,y in zip(a,b))
with zipfile.ZipFile(NATIVE/'gradient.pptx') as z:
 for i,(angle,scaled) in enumerate([(0,0),(5400000,0),(2700000,0),(2700000,1),(0,0),(16200000,0)],1):
  root=ET.fromstring(z.read(f'ppt/slides/slide{i}.xml')); lin=root.find('p:cSld/p:bg/p:bgPr/a:gradFill/a:lin',NS)
  assert lin is not None and int(lin.get('ang'))==angle and int(lin.get('scaled'))==scaled
  assert len(root.findall('p:cSld/p:bg/p:bgPr/a:gradFill/a:gsLst/a:gs',NS))==2
points=[(720,135),(240,405),(480,270)]
observations={}
for i in range(1,7):
 native=Image.open(NATIVE/f'native-{i}.png').convert('RGB'); svg=Image.open(PREVIEW/f'svg-{i}.png').convert('RGB')
 assert native.size==(1920,1080) and svg.size==(960,540)
 observations[str(i)]={'native':[native.getpixel((2*x,2*y)) for x,y in points],'svg':[svg.getpixel((x,y)) for x,y in points]}
assert diff(*observations['4']['native'][:2])<=4
assert diff(*observations['4']['svg'][:2])<=4
assert diff(*observations['3']['native'][:2])>=20
assert diff(observations['1']['native'][2],observations['1']['svg'][2])>=30
assert diff(observations['1']['native'][2],observations['5']['native'][2])>=20
out={'scope':'PowerPoint 16.104 first-open slideshow; authored two-stop background at 0, 90, 45 scaled and unscaled, half-alpha and 270 degrees. Screenshots alone do not prove UI provenance. No native edit or save/reopen. SVG color differs; this is not general fidelity evidence.', 'nativePptxSha256':digest(NATIVE/'gradient.pptx'),'producerSha256':digest(ROOT/'scripts/prepare-background-gradient-native.mjs'),'oracleSha256':digest(Path(__file__)),'nativeScreenshots':[digest(NATIVE/f'native-{i}.png') for i in range(1,7)],'svgScreenshots':[digest(PREVIEW/f'svg-{i}.png') for i in range(1,7)],'samplePoints':points,'samples':observations}
(ROOT/'docs/testing/background-gradient-native-evidence.json').write_text(json.dumps(out,indent=2)+'\n')
print(json.dumps(observations,indent=2))
