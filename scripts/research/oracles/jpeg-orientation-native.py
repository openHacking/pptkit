#!/usr/bin/env python3
"""Verify retained pixels and package; does not witness manual native UI provenance."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import base64
import json
import posixpath
import xml.etree.ElementTree as ET
from PIL import Image, ImageOps, ImageChops
ROOT=Path(__file__).resolve().parents[3]; BASE=ROOT/'.tmp/jpeg-orientation-native-01'
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','rel':'http://schemas.openxmlformats.org/package/2006/relationships'}
def check(value,message):
    if not value: raise ValueError(message)
def sha(b): return sha256(b).hexdigest()
def inspect(jpeg,png,pptx,svg,native,browser):
    im=Image.open(BytesIO(jpeg));check(im.size==(300,200) and im.getexif().get(274)==6,'JPEG source metadata')
    reference=Image.open(BytesIO(png)).convert('RGB'); normalized=ImageOps.exif_transpose(im).convert('RGB')
    check(reference.size==(200,300) and ImageChops.difference(reference,normalized).getbbox() is None,'control is not decoded oriented JPEG')
    with ZipFile(BytesIO(pptx)) as z:
        check(len(z.namelist())==len(set(z.namelist())),'duplicate package parts')
        slide=ET.fromstring(z.read('ppt/slides/slide1.xml')); rel=ET.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))
        refs={r.get('Id'):r.get('Target') for r in rel.findall('rel:Relationship',N) if r.get('Type','').endswith('/image') and r.get('TargetMode') is None}
        pictures=slide.findall('.//p:pic',N);check(len(pictures)==2,'picture count')
        for picture,expected,name,x in zip(pictures,[jpeg,png],['image-exif','image-normalized'],[1270000,6350000]):
            nv=picture.find('p:nvPicPr/p:cNvPr',N);check(nv.get('name')==name,'picture identity')
            blip=picture.find('p:blipFill/a:blip',N);ref=blip.get('{%s}embed'%N['r']);check(ref in refs,'relationship')
            target=posixpath.normpath(posixpath.join('ppt/slides',refs[ref]));check(target.startswith('ppt/media/') and z.read(target)==expected,'media payload')
            transform=picture.find('p:spPr/a:xfrm',N);off=transform.find('a:off',N);ext=transform.find('a:ext',N)
            check([int(off.get('x')),int(off.get('y')),int(ext.get('cx')),int(ext.get('cy'))]==[x,1524000,2540000,3810000],'picture box')
    root=ET.fromstring(svg);images=root.findall('.//{http://www.w3.org/2000/svg}image');check(len(images)==2,'preview count')
    for item,expected in zip(images,[jpeg,png]):
        href=item.get('href','');check(href.startswith('data:image/') and base64.b64decode(href.split(',',1)[1],validate=True)==expected,'preview bytes')
    def classify(pixel):
        rgb=tuple(1 if value>180 else 0 if value<80 else -1 for value in pixel)
        names={(1,0,0):'red',(0,1,0):'green',(0,0,1):'blue',(1,1,0):'yellow',(1,0,1):'magenta',(0,1,1):'cyan'}
        return names.get(rgb,'other')
    native=Image.open(BytesIO(native)).convert('RGB');browser=Image.open(BytesIO(browser)).convert('RGB')
    check(native.size==(1920,1080) and browser.size==(960,540),'screenshot size')
    def grid(image,xs,ys):return [[classify(image.getpixel((x,y))) for x in xs] for y in ys]
    raw=[['red','green','blue'],['yellow','magenta','cyan']]
    rotated=[['yellow','red'],['magenta','green'],['cyan','blue']]
    nativeLeft=grid(native,[266,400,533],[390,690]);nativeRight=grid(native,[1100,1300],[340,540,740])
    browserLeft=grid(browser,[150,250],[170,270,370]);browserRight=grid(browser,[550,650],[170,270,370])
    check(nativeLeft==raw and nativeRight==rotated,'native grids')
    check(browserLeft==rotated and browserRight==rotated,'browser grids')
    return {'nativeJpeg':nativeLeft,'nativeControl':nativeRight,'browserJpeg':browserLeft,'browserControl':browserRight}
def main():
    names=['orientation6.jpg','normalized.png','candidate.pptx','preview.svg','powerpoint.png','browser.png'];data=[(BASE/n).read_bytes() for n in names]
    producer=json.loads((BASE/'producer.json').read_text()); check([sha(b) for b in data[:4]]==[producer[k] for k in ['jpegSha256','pngSha256','pptxSha256','svgSha256']],'producer hashes')
    browserInfo=json.loads((BASE/'browser.json').read_text());check(browserInfo['requests']==[],'browser network')
    result=inspect(*data); rejected=0
    for corrupt in [(data[1],data[1],*data[2:]),(data[0],data[0],*data[2:]),(*data[:4],data[5],data[5])]:
        try: inspect(*corrupt)
        except Exception: rejected+=1
    check(rejected==3,'negative cases')
    scripts=['scripts/prepare-jpeg-orientation-source.py','scripts/prepare-jpeg-orientation-native.mjs','scripts/inspect-jpeg-orientation-browser.mjs']
    evidence={'scope':'One synthetic JPEG and PNG control, PowerPoint 16.104 manual slideshow and Chromium browser preview. Pixel samples and package inspection do not prove UI provenance, full visual equivalence or other images/versions.',
              'directory':str(BASE.relative_to(ROOT)),'files':{name:sha(data) for name,data in zip(names,data)},'producerSha256':sha((BASE/'producer.json').read_bytes()),'browserInfo':browserInfo,'observedGrids':result,'negativeCasesRejected':rejected,'oracleSha256':sha(Path(__file__).read_bytes()),'scripts':{name:sha((ROOT/name).read_bytes()) for name in scripts}}
    (ROOT/'docs/testing/jpeg-orientation-native-evidence.json').write_text(json.dumps(evidence,indent=2)+'\n');print(json.dumps(result))
if __name__=='__main__':main()
