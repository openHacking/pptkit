#!/usr/bin/env python3
"""Check normalized sample files and pixels; native UI actions remain observations."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import json
import posixpath
import xml.etree.ElementTree as ET
from PIL import Image,ImageOps,ImageChops
ROOT=Path(__file__).resolve().parents[3]; BASE=ROOT/'.tmp/jpeg-orientation-native-01'
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','rel':'http://schemas.openxmlformats.org/package/2006/relationships'}
def check(value,message):
    if not value: raise ValueError(message)
def sha(data):return sha256(data).hexdigest()
def inspect(jpeg,png,pptx,screenshot):
    source=Image.open(BytesIO(jpeg)); check(source.getexif().get(274)==6,'source orientation')
    expected=ImageOps.exif_transpose(source).convert('RGB');control=Image.open(BytesIO(png)).convert('RGB')
    check(expected.size==control.size and ImageChops.difference(expected,control).getbbox() is None,'control pixel normalization')
    with ZipFile(BytesIO(pptx)) as z:
        check(len(z.namelist())==len(set(z.namelist())),'duplicate package')
        slide=ET.fromstring(z.read('ppt/slides/slide1.xml'));rels=ET.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))
        ids={r.get('Id'):r.get('Target') for r in rels.findall('rel:Relationship',N) if r.get('Type','').endswith('/image') and r.get('TargetMode') is None}
        pictures=slide.findall('.//p:pic',N);check(len(pictures)==2,'picture count');targets=[]
        for picture,name,x in zip(pictures,['image-exif','image-normalized'],[1270000,6350000]):
            nv=picture.find('p:nvPicPr/p:cNvPr',N);check(nv.get('name')==name,'picture name')
            blip=picture.find('p:blipFill/a:blip',N);ref=blip.get('{%s}embed'%N['r']);check(ref in ids,'image reference')
            target=posixpath.normpath(posixpath.join('ppt/slides',ids[ref]));check(target.startswith('ppt/media/') and z.read(target)==png,'normalized payload');targets.append(target)
            transform=picture.find('p:spPr/a:xfrm',N);off=transform.find('a:off',N);ext=transform.find('a:ext',N)
            check([int(off.get('x')),int(off.get('y')),int(ext.get('cx')),int(ext.get('cy'))]==[x,1524000,2540000,3810000],'picture box')
    screen=Image.open(BytesIO(screenshot)).convert('RGB');check(screen.size==(1920,1080),'screen size')
    def classify(pixel):
        vals=tuple(1 if x>180 else 0 if x<80 else -1 for x in pixel)
        return {(1,0,0):'red',(0,1,0):'green',(0,0,1):'blue',(1,1,0):'yellow',(1,0,1):'magenta',(0,1,1):'cyan'}.get(vals,'other')
    def grid(xs):return [[classify(screen.getpixel((x,y))) for x in xs] for y in [340,540,740]]
    left,right=grid([300,500]),grid([1100,1300]);rotated=[['yellow','red'],['magenta','green'],['cyan','blue']]
    check(left==right==rotated,'screen grid');return {'pictureTargets':targets,'left':left,'right':right}
def main():
    names=['orientation6.jpg','normalized.png','normalized-candidate.pptx','normalized-powerpoint.png'];data=[(BASE/n).read_bytes() for n in names]
    original=json.loads((ROOT/'docs/testing/jpeg-orientation-native-evidence.json').read_text()); report=json.loads((BASE/'normalized-producer.json').read_text())
    check(sha(data[0])==original['files']['orientation6.jpg'] and sha(data[1])==original['files']['normalized.png'],'source hashes')
    check(sha(data[2])==report['pptxSha256'] and sha(data[1])==report['normalizedPngSha256'] and sha(data[0])==report['sourceJpegSha256'],'producer hashes')
    check([w['code'] for w in report['warnings']]==['asset-orientation-normalized'],'conversion report')
    result=inspect(*data);bad=0
    for sample in [(data[0],data[0],data[2],data[3]), (data[0],data[1],data[2],(BASE/'powerpoint.png').read_bytes()),(data[0],data[1],(BASE/'candidate.pptx').read_bytes(),data[3])]:
        try:inspect(*sample)
        except Exception:bad+=1
    check(bad==3,'negative controls')
    evidence={'scope':'One resolver-attested PNG normalization of an EXIF 6 JPEG, with PowerPoint 16.104 manual slideshow. Pixel samples do not prove UI provenance or general crop/color fidelity.',
      'directory':str(BASE.relative_to(ROOT)),'files':{name:sha(b) for name,b in zip(names,data)},'producerSha256':sha((BASE/'normalized-producer.json').read_bytes()),'originalEvidenceSha256':sha((ROOT/'docs/testing/jpeg-orientation-native-evidence.json').read_bytes()),
      'result':result,'negativeCasesRejected':bad,'oracleSha256':sha(Path(__file__).read_bytes()),'producerScriptSha256':sha((ROOT/'scripts/prepare-jpeg-orientation-normalized.mjs').read_bytes())}
    (ROOT/'docs/testing/jpeg-orientation-normalized-evidence.json').write_text(json.dumps(evidence,indent=2)+'\n');print(json.dumps(result))
if __name__=='__main__':main()
