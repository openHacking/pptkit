#!/usr/bin/env python3
"""Verify one native crop-handle edit and imported SVG preview."""
import json, zipfile, xml.etree.ElementTree as ET
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from PIL import Image
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/".tmp/jpeg-orientation-native-01"
NS={"p":"http://schemas.openxmlformats.org/presentationml/2006/main","a":"http://schemas.openxmlformats.org/drawingml/2006/main","r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships","rel":"http://schemas.openxmlformats.org/package/2006/relationships"}
def require(condition,message):
    if not condition: raise ValueError(message)
def digest(data): return sha256(data).hexdigest()
def inspect(package,native,preview,normalized,control):
    with zipfile.ZipFile(BytesIO(package)) as z:
        slide=ET.fromstring(z.read("ppt/slides/slide1.xml"));rels=ET.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels"))
        targets={r.get("Id"):r.get("Target") for r in rels.findall("rel:Relationship",NS)}
        pics={p.find("p:nvPicPr/p:cNvPr",NS).get("name"):p for p in slide.findall(".//p:pic",NS)};require(set(pics)=={"mapped","control"},"picture identities")
        for name,data,x,w in [("mapped",normalized,"2295938","2276061"),("control",control,"7620000","3048000")]:
            pic=pics[name];blip=pic.find("p:blipFill/a:blip",NS);target=targets[blip.get("{%s}embed"%NS["r"])];require(target.startswith("../media/") and z.read("ppt/"+target[3:])==data,"media")
            transform=pic.find("p:spPr/a:xfrm",NS);require(transform.find("a:off",NS).attrib=={"x":x,"y":"1651000"} and transform.find("a:ext",NS).attrib=={"cx":w,"cy":"3048000"},"frame")
        crop=pics["mapped"].find("p:blipFill/a:srcRect",NS);require(crop is not None and {s:int(crop.get(s,"0")) for s in "ltrb"}=={"l":25326,"t":0,"r":0,"b":33333},"edited crop")
        require(pics["control"].find("p:blipFill/a:srcRect",NS) is None,"control crop")
    results=[]
    for encoded in (native,preview):
        with Image.open(BytesIO(encoded)) as image:
            im=image.convert("RGB");require(im.size==(1920,1080),"image size")
            def name(x,y):
                r,g,b=im.getpixel((x,y))
                if r>200 and g>200 and b<80:return "yellow"
                if r>200 and g<80 and b<80:return "red"
                if r>200 and g<80 and b>200:return "magenta"
                if r<80 and g>200 and b<80:return "green"
                if min(r,g,b)>240:return "white"
                return "other"
            points=[(340,380,"white"),(400,380,"yellow"),(550,380,"red"),(400,620,"magenta"),(550,620,"green"),(740,380,"white"),(1320,380,"yellow"),(1560,380,"red"),(1320,620,"magenta"),(1560,620,"green")]
            require(all(name(x,y)==expected for x,y,expected in points),"sampled pixels")
            row=[x for x in range(300,800) if name(x,380) in ("yellow","red")];require(row and 360<=min(row)<=363 and max(row)==719,"edited image span")
            results.append([min(row),max(row)])
    return results
def main():
    names=["crop-native-edit.pptx","crop-native-edit.png","crop-native-edit-preview.png","browser-normalized.png","crop-control.png"]
    package,native,preview,normalized,control=[(BASE/name).read_bytes() for name in names]
    report_bytes=(BASE/"crop-native-edit-report.json").read_bytes();report=json.loads(report_bytes)
    require(report["inputSha256"]==digest(package) and report["previewSha256"]==digest(preview) and report["svgSha256"]==digest((BASE/"crop-native-edit.svg").read_bytes()) and report["requests"]==[] and report["noOpByteIdentical"] is True,"report bindings")
    observed=inspect(package,native,preview,normalized,control)
    cases=[((BASE/"crop-native-round1.pptx").read_bytes(),native,preview), (package,(BASE/"crop-native-round1.png").read_bytes(),preview), (package,native,(BASE/"crop-native-round1.png").read_bytes())]
    rejected=0
    for args in cases:
        try: inspect(*args,normalized,control)
        except Exception: rejected+=1
    require(rejected==3,"negative controls")
    result={"scope":"PowerPoint 16.104 manual left crop-handle drag, save, close, reopen and slideshow; Chromium SVG preview sampled on one orientation-6 image", "directory":str(BASE.relative_to(ROOT)),"files":{n:digest((BASE/n).read_bytes()) for n in names},"priorEvidenceSha256":digest((ROOT/"docs/testing/jpeg-crop-reopen-evidence.json").read_bytes()),"reportSha256":digest(report_bytes),"scriptSha256":digest((ROOT/"scripts/inspect-jpeg-crop-edit.mjs").read_bytes()),"oracleSha256":digest(Path(__file__).read_bytes()),"imageHorizontalSpans":observed,"negativeCasesRejected":rejected}
    (ROOT/"docs/testing/jpeg-crop-edit-evidence.json").write_text(json.dumps(result,indent=2)+"\n");print(json.dumps({"spans":observed,"negativeCasesRejected":rejected}))
if __name__=="__main__": main()
