#!/usr/bin/env python3
"""Check a named EXIF-6 crop against independent decoded pixels and native screenshot."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import json, zipfile, xml.etree.ElementTree as ET
from PIL import Image, ImageChops, ImageOps
ROOT = Path(__file__).resolve().parents[3]
BASE = ROOT / ".tmp/jpeg-orientation-native-01"
NS = {"p":"http://schemas.openxmlformats.org/presentationml/2006/main", "a":"http://schemas.openxmlformats.org/drawingml/2006/main", "r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rel":"http://schemas.openxmlformats.org/package/2006/relationships"}
def require(condition, text):
    if not condition: raise ValueError(text)
def digest(data): return sha256(data).hexdigest()
def inspect(source, normalized, control, package, screenshot):
    with Image.open(BytesIO(source)) as raw:
        require(raw.size == (300,200) and raw.getexif().get(274) == 6, "source orientation")
        expected = raw.convert("RGB").crop((0,0,200,200)).transpose(Image.Transpose.ROTATE_270)
        all_oriented = ImageOps.exif_transpose(raw).convert("RGB")
    with Image.open(BytesIO(normalized)) as image:
        require(image.size == (200,300) and ImageChops.difference(image.convert("RGB"),all_oriented).getbbox() is None,"normalized pixels")
    with Image.open(BytesIO(control)) as image:
        require(image.size == (200,200) and ImageChops.difference(image.convert("RGB"),expected).getbbox() is None,"independent physical crop")
    with zipfile.ZipFile(BytesIO(package)) as archive:
        slide=ET.fromstring(archive.read("ppt/slides/slide1.xml"));rels=ET.fromstring(archive.read("ppt/slides/_rels/slide1.xml.rels"))
        targets={r.attrib["Id"]:r.attrib["Target"] for r in rels.findall("rel:Relationship",NS)}
        pictures={p.find("p:nvPicPr/p:cNvPr",NS).attrib["name"]:p for p in slide.findall(".//p:pic",NS)}
        require(set(pictures)=={"mapped","control"},"picture identities")
        for name,data,x in [("mapped",normalized,"1524000"),("control",control,"7620000")]:
            pic=pictures[name];embed=pic.find("p:blipFill/a:blip",NS).attrib["{%s}embed"%NS["r"]]
            target=targets[embed];require(target.startswith("../media/"),"relationship target")
            require(archive.read("ppt/"+target[3:])==data,"image bytes")
            transform=pic.find("p:spPr/a:xfrm",NS)
            require(transform.find("a:off",NS).attrib=={"x":x,"y":"1651000"} and transform.find("a:ext",NS).attrib=={"cx":"3048000","cy":"3048000"},"box")
        require(pictures["mapped"].find("p:blipFill/a:srcRect",NS).attrib=={"l":"0","t":"0","r":"0","b":"33333"},"crop")
        require(pictures["control"].find("p:blipFill/a:srcRect",NS) is None,"control must be uncropped")
    with Image.open(BytesIO(screenshot)) as image:
        image=image.convert("RGB");require(image.size==(1920,1080),"screenshot dimensions")
        def classify(point):
            r,g,b=image.getpixel(point)
            if r>200 and g>200 and b<80:return "yellow"
            if r>200 and g<80 and b<80:return "red"
            if r>200 and g<80 and b>200:return "magenta"
            if r<80 and g>200 and b<80:return "green"
            return "unexpected"
        observed=[[classify((x,y)) for y in (380,620) for x in (left+120,left+360)] for left in (240,1200)]
        require(observed==[["yellow","red","magenta","green"]]*2,"screenshot color cells")
    return observed
def main():
    names=["orientation6.jpg","browser-normalized.png","crop-control.png","crop-candidate.pptx","crop-powerpoint.png"]
    source,normalized,control,package,screenshot=[(BASE/name).read_bytes() for name in names]
    producer=json.loads((BASE/"crop-producer.json").read_text())
    require(producer["pptxSha256"]==digest(package) and producer["normalizedSha256"]==digest(normalized) and producer["controlSha256"]==digest(control) and producer["warnings"]==[],"producer hashes")
    observed=inspect(source,normalized,control,package,screenshot)
    controls=[(source,source,control,package,screenshot),(source,normalized,normalized,package,screenshot),(source,normalized,control,(BASE/"browser-normalized-candidate.pptx").read_bytes(),screenshot),(source,normalized,control,package,(BASE/"powerpoint.png").read_bytes())]
    rejected=0
    for item in controls:
        try: inspect(*item)
        except Exception: rejected+=1
    require(rejected==len(controls),"negative controls")
    paths=["scripts/prepare-jpeg-crop-control.py","scripts/prepare-jpeg-crop-native.mjs","packages/pptx-exporter/src/assets/orient-jpeg-crop.ts"]
    evidence={"scope":"PowerPoint 16.104 manual slideshow; one EXIF 6 crop and four sampled cells per image; no save/reopen claim", "directory":str(BASE.relative_to(ROOT)),"files":{name:digest((BASE/name).read_bytes()) for name in names},"producerSha256":digest((BASE/"crop-producer.json").read_bytes()),"oracleSha256":digest(Path(__file__).read_bytes()),"scripts":{path:digest((ROOT/path).read_bytes()) for path in paths},"observed":observed,"negativeCasesRejected":rejected}
    (ROOT/"docs/testing/jpeg-crop-native-evidence.json").write_text(json.dumps(evidence,indent=2)+"\n")
    print(json.dumps({"observed":observed,"negativeCasesRejected":rejected}))
if __name__=="__main__": main()
