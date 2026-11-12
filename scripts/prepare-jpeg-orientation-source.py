#!/usr/bin/env python3
"""Generate an asymmetric orientation-6 JPEG and a decoded, normalized PNG control."""
from pathlib import Path
import sys
from PIL import Image, ImageDraw, ImageOps
if len(sys.argv) != 2: raise SystemExit('Usage: python3 scripts/prepare-jpeg-orientation-source.py <new-directory>')
directory=Path(sys.argv[1]); directory.mkdir(parents=True,exist_ok=False)
im=Image.new('RGB',(300,200)); draw=ImageDraw.Draw(im)
for index,color in enumerate(('red','lime','blue','yellow','magenta','cyan')):
    x=(index%3)*100;y=(index//3)*100;draw.rectangle((x,y,x+99,y+99),fill=color)
exif=Image.Exif();exif[274]=6;im.save(directory/'orientation6.jpg','JPEG',quality=100,subsampling=0,exif=exif)
with Image.open(directory/'orientation6.jpg') as image: ImageOps.exif_transpose(image).save(directory/'normalized.png')
