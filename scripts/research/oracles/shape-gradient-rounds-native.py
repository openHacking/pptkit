#!/usr/bin/env python3
"""Verify three named PowerPoint save/reopen rounds for one edited shape fixture."""
import hashlib,json,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
from PIL import Image
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/'.tmp/shape-gradient-native-01';N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def inspect(path):
 values=[]
 with zipfile.ZipFile(path) as z:
  for i in range(1,4):
   root=ET.fromstring(z.read(f'ppt/slides/slide{i}.xml'));shapes=root.findall('.//p:sp',N);assert len(shapes)==1;sh=shapes[0];grad=sh.find('p:spPr/a:gradFill',N);lin=grad.find('a:lin',N);stops=[]
   for stop in grad.findall('a:gsLst/a:gs',N):
    color=stop.find('a:srgbClr',N);alpha=color.find('a:alpha',N);stops.append([stop.get('pos'),color.get('val'),alpha.get('val') if alpha is not None else '100000'])
   shape=sh.find('p:spPr/a:xfrm',N);group=root.find('.//p:grpSp/p:grpSpPr/a:xfrm',N);values.append({'angle':lin.get('ang'),'scaled':lin.get('scaled'),'stops':stops,'shapeRotation':shape.get('rot','0'),'groupRotation':group.get('rot','0') if group is not None else None})
 return values
files=[BASE/f'shape-gradient-{x}.pptx' for x in ('edited','round2','round3')];actual=[inspect(p) for p in files];assert actual[0]==actual[1]==actual[2];assert actual[0][0]['angle']=='7200000' and actual[0][0]['scaled']=='0' and actual[0][0]['stops'][0][2]=='50000';assert actual[0][1]['shapeRotation']=='1800000' and actual[0][2]['groupRotation']=='1800000'
points=[(480,270),(300,180),(660,360),(300,270),(660,270),(480,180),(480,360),(100,100)];samples={}
for cycle in (2,3):
 samples[str(cycle)]=[]
 for page in (1,2,3):
  image=Image.open(BASE/f'round{cycle}-page{page}.png').convert('RGB');assert image.size==(1920,1080);samples[str(cycle)].append([image.getpixel((2*x,2*y)) for x,y in points])
assert samples['2']==samples['3'];assert samples['2'][0]!=samples['2'][1] and samples['2'][1]!=samples['2'][2]
result={'scope':'PowerPoint 16.104: one native shape edit, then three named save/close/reopen rounds; XML sampled on all rounds, slide-show pixel samples on rounds two and three. Screenshots and hashes do not independently prove UI provenance. No general gradient or source-preserving edit claim.','files':[sha(p) for p in files],'screenshots':{str(i):[sha(BASE/f'round{i}-page{j}.png') for j in (1,2,3)] for i in (2,3)},'oracleSha256':sha(Path(__file__)),'values':actual,'samplePoints':points,'samples':samples}
(ROOT/'docs/testing/shape-gradient-native-rounds-evidence.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps({'files':result['files'],'values':actual[0]}))
