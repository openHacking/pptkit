#!/usr/bin/env python3
"""Bounded point-sampling and XML check of one PowerPoint save/reopen observation."""
import hashlib
import json
from math import cos, sin, pi, hypot
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile
from PIL import Image

ROOT=Path(__file__).resolve().parents[2]
DIR=ROOT/'.tmp/native-parallelogram-slant'
RECORD=ROOT/'docs/testing/parallelogram-slant-powerpoint-evidence.json'
FILES=('source.pptx','native-round-1.pptx','first.ax.txt','first.jpg','reopen.ax.txt','reopen.jpg')
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
SHAPES=[('zero',30,30,220,80,0,False,0),('mid',310,30,80,220,0,False,40000),('max-rotated',40,190,220,80,30,True,100000)]

def polygon(shape):
    _,x,y,w,h,rotation,flip,slant=shape
    inset=min(w,h)*slant/100000
    points=[(x,y+h),(x+inset,y),(x+w,y),(x+w-inset,y+h)]
    cx,cy=x+w/2,y+h/2
    c,s=cos(rotation*pi/180),sin(rotation*pi/180)
    return [(cx+(px-cx)*(-1 if flip else 1)*c-(py-cy)*s,cy+(px-cx)*(-1 if flip else 1)*s+(py-cy)*c) for px,py in points]

def segments(poly): return zip(poly,poly[1:]+poly[:1])
def inside(point,poly):
    x,y=point; result=False
    for (x1,y1),(x2,y2) in segments(poly):
        if (y1>y)!=(y2>y) and x<(x2-x1)*(y-y1)/(y2-y1)+x1: result=not result
    return result

def distance(point,a,b):
    x,y=point;dx=b[0]-a[0];dy=b[1]-a[1]
    t=max(0,min(1,((x-a[0])*dx+(y-a[1])*dy)/(dx*dx+dy*dy)))
    return hypot(x-a[0]-t*dx,y-a[1]-t*dy)

def check():
    hashes={f:hashlib.sha256((DIR/f).read_bytes()).hexdigest() for f in FILES}
    assert hashlib.sha256((ROOT/'tests/fixtures/parallelogram-slant-powerpoint-resave.pptx').read_bytes()).hexdigest()==hashes['native-round-1.pptx']
    for name in FILES[:2]:
        with zipfile.ZipFile(DIR/name) as package: root=ET.fromstring(package.read('ppt/slides/slide1.xml'))
        shapes=root.findall('.//p:sp',NS);assert len(shapes)==3
        for node,(identity,x,y,w,h,rotation,flip,slant) in zip(shapes,SHAPES):
            assert node.find('p:nvSpPr/p:cNvPr',NS).get('name')==identity
            props=node.find('p:spPr',NS);geom=props.find('a:prstGeom',NS)
            assert geom.get('prst')=='parallelogram'
            guides=list(geom.find('a:avLst',NS));assert len(guides)==1 and guides[0].get('name')=='adj' and guides[0].get('fmla')==f'val {slant}'
            trans=props.find('a:xfrm',NS);off=trans.find('a:off',NS);ext=trans.find('a:ext',NS)
            assert [int(off.get('x')),int(off.get('y')),int(ext.get('cx')),int(ext.get('cy'))]==[v*12700 for v in (x,y,w,h)]
            assert int(trans.get('rot','0'))==rotation*60000 and trans.get('flipH','0')==str(int(flip))
            assert props.find('a:solidFill/a:srgbClr',NS).get('val')=='FF0000'
    polys=[polygon(shape) for shape in SHAPES];samples={}
    for stem,deck in [('first','source.pptx'),('reopen','native-round-1.pptx')]:
        ax=(DIR/f'{stem}.ax.txt').read_text();assert f'/native-parallelogram-slant/{deck}' in ax
        assert all(f'形状 (disabled) {shape[0]}\n' in ax for shape in SHAPES)
        image=Image.open(DIR/f'{stem}.jpg').convert('RGB');assert image.size==(1600,900)
        counts=[0,0]
        for y in range(10,351,8):
            for x in range(10,591,8):
                # The captured pointer and its halo obscure this blank area.
                if 485 <= x <= 515 and 180 <= y <= 215: continue
                if min(distance((x,y),a,b) for poly in polys for a,b in segments(poly))<4:continue
                expected=any(inside((x,y),poly) for poly in polys)
                r,g,b=image.getpixel((round(350+x*1.92),round(158+y*1.92)))
                actual=r>220 and g<60 and b<60
                assert expected==actual,(stem,x,y,(r,g,b),expected)
                if not expected: assert min(r,g,b)>220,(stem,x,y,(r,g,b),"expected white")
                counts[int(expected)]+=1
        assert counts[0]>2000 and counts[1]>500;samples[stem]={'outside':counts[0],'inside':counts[1],'mismatches':0}
    return {'scope':'One opaque red slide with slants 0, 0.4 and 1 including a rotation/flip, one PowerPoint save and reopen. Fixed grid points at least 4 pt from edges, excluding the pointer halo, are compared against an independent polygon model; native UI provenance is observed, not authenticated. Edges, text, other adjustments, interactive edits and other versions remain unverified.','reopenRounds':1,'samples':samples,'artifactDirectory':'.tmp/native-parallelogram-slant','sha256':hashes}

result=check()
if sys.argv[1:]==['--record']: RECORD.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n')
else: assert json.loads(RECORD.read_text())==result
print(json.dumps({'verified':True,'samplesPerCapture':sum(result['samples']['first'][k] for k in ('outside','inside')),'reopenRounds':1}))
