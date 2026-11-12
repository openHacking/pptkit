#!/usr/bin/env python3
"""Build the manually assembled placeholder research file from the slant sample."""
from pathlib import Path
import re
import zipfile

root=Path(__file__).resolve().parents[2]
source=root/'.tmp/native-parallelogram-slant/source.pptx'
output=root/'.tmp/native-inherited-slant/source.pptx'
with zipfile.ZipFile(source) as z: parts={name:z.read(name) for name in z.namelist()}
path='ppt/slides/slide1.xml';slide=parts[path].decode();shapes=re.findall(r'<p:sp>.*?</p:sp>',slide);assert len(shapes)==3
geom=lambda s:re.sub(r'<a:prstGeom .*?</a:prstGeom>','',s,count=1)
ph=lambda s,role,index:s.replace('<p:nvPr></p:nvPr>',f'<p:nvPr><p:ph type="{role}" idx="{index}"/></p:nvPr>',1)
assert 'fmla="val 0"' in shapes[0] and 'fmla="val 40000"' in shapes[1]
layout1=ph(shapes[0],'body',1).replace('name="zero"','name="layout-body"',1).replace('fmla="val 0"','fmla="val 40000"',1)
layout2=geom(ph(shapes[1],'title',2)).replace('name="mid"','name="layout-title"',1)
master=ph(shapes[1],'title',2).replace('name="mid"','name="master-title"',1).replace('fmla="val 40000"','fmla="val 80000"',1)
slide=slide.replace(shapes[0],geom(ph(shapes[0],'body',1)),1).replace(shapes[1],geom(ph(shapes[1],'title',2)),1)
parts[path]=slide.encode()
for name,nodes in [('ppt/slideLayouts/slideLayout1.xml',layout1+layout2),('ppt/slideMasters/slideMaster1.xml',master)]:
    xml=parts[name].decode();assert xml.count('</p:spTree>')==1;parts[name]=xml.replace('</p:spTree>',nodes+'</p:spTree>').encode()
output.parent.mkdir(parents=True,exist_ok=True)
with zipfile.ZipFile(output,'x',zipfile.ZIP_DEFLATED) as z:
    for name,data in parts.items():z.writestr(name,data)
