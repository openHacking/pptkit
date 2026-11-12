#!/usr/bin/env python3
"""Check one manual native edit/save/reopen without relying on XML byte identity."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import importlib.util, json, zipfile, xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[3]; BASE=ROOT/".tmp/jpeg-orientation-native-01"
path=ROOT/"scripts/research/oracles/jpeg-crop-native.py"
spec=importlib.util.spec_from_file_location("reference",path); reference=importlib.util.module_from_spec(spec); spec.loader.exec_module(reference)
NS=reference.NS;require=reference.require;digest=reference.digest

def inspect(package, screenshot, normalized, control):
    with zipfile.ZipFile(BytesIO(package)) as archive:
        slide=ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        require("".join(t.text or "" for t in slide.findall(".//a:t",NS)).endswith(" R1"),"native edit marker")
        rels=ET.fromstring(archive.read("ppt/slides/_rels/slide1.xml.rels"))
        targets={r.attrib["Id"]:r.attrib["Target"] for r in rels.findall("rel:Relationship",NS)}
        pictures={p.find("p:nvPicPr/p:cNvPr",NS).attrib["name"]:p for p in slide.findall(".//p:pic",NS)}
        require(set(pictures)=={"mapped","control"},"picture identities")
        for name,data,x in [("mapped",normalized,"1524000"),("control",control,"7620000")]:
            pic=pictures[name];embed=pic.find("p:blipFill/a:blip",NS).attrib["{%s}embed"%NS["r"]]
            target=targets[embed];require(target.startswith("../media/"),"relationship target")
            require(archive.read("ppt/"+target[3:])==data,"media bytes")
            xfrm=pic.find("p:spPr/a:xfrm",NS)
            require(xfrm.find("a:off",NS).attrib=={"x":x,"y":"1651000"} and xfrm.find("a:ext",NS).attrib=={"cx":"3048000","cy":"3048000"},"frame")
        crop=pictures["mapped"].find("p:blipFill/a:srcRect",NS)
        require(crop is not None and {side:int(crop.get(side,"0")) for side in "ltrb"}=={"l":0,"t":0,"r":0,"b":33333},"crop semantics")
        require(pictures["control"].find("p:blipFill/a:srcRect",NS) is None,"control crop")
    # The original oracle performs source/Pillow checks and screenshot sampling on the original package.
    observed=reference.inspect((BASE/"orientation6.jpg").read_bytes(),normalized,control,(BASE/"crop-candidate.pptx").read_bytes(),screenshot)
    return observed

def replace_slide(package, old, new):
    output=BytesIO()
    with zipfile.ZipFile(BytesIO(package)) as src, zipfile.ZipFile(output,"w") as dst:
        for item in src.infolist():
            data=src.read(item.filename)
            if item.filename=="ppt/slides/slide1.xml":
                require(old in data,"negative-control source");data=data.replace(old,new,1)
            dst.writestr(item,data)
    return output.getvalue()

def main():
    package=(BASE/"crop-native-round1.pptx").read_bytes(); screenshot=(BASE/"crop-native-round1.png").read_bytes()
    normalized=(BASE/"browser-normalized.png").read_bytes(); control=(BASE/"crop-control.png").read_bytes()
    require(digest(package)!=digest((BASE/"crop-candidate.pptx").read_bytes()),"package rewritten")
    observed=inspect(package,screenshot,normalized,control)
    parser_bytes=(BASE/"crop-native-parser.json").read_bytes(); parser=json.loads(parser_bytes)
    require(parser["inputSha256"]==digest(package) and parser["noOpByteIdentical"] is True and parser["mappedCrop"]=={"left":0,"top":0,"right":0,"bottom":0.33333},"parser report")
    cases=[((BASE/"crop-candidate.pptx").read_bytes(),screenshot),
           (replace_slide(package,b'<a:srcRect b="33333"',b'<a:srcRect b="0"'),screenshot),
           (package,(BASE/"powerpoint.png").read_bytes())]
    rejected=0
    for item in cases:
        try: inspect(*item,normalized,control)
        except Exception: rejected+=1
    require(rejected==len(cases),"negative controls")
    result={"scope":"PowerPoint 16.104 manual title edit, save, close, reopen and slideshow for one EXIF 6 crop; screenshot sampled, not independent UI witness", "directory":str(BASE.relative_to(ROOT)), "sourceEvidenceSha256":digest((ROOT/"docs/testing/jpeg-crop-native-evidence.json").read_bytes()),"round1PptxSha256":digest(package),"round1ScreenshotSha256":digest(screenshot),"referenceOracleSha256":digest(path.read_bytes()),"oracleSha256":digest(Path(__file__).read_bytes()),"parserReportSha256":digest(parser_bytes),"parserScriptSha256":digest((ROOT/"scripts/inspect-jpeg-crop-reopen.mjs").read_bytes()),"observed":observed,"negativeCasesRejected":rejected}
    (ROOT/"docs/testing/jpeg-crop-reopen-evidence.json").write_text(json.dumps(result,indent=2)+"\n");print(json.dumps({"observed":observed,"negativeCasesRejected":rejected}))
if __name__=="__main__":main()
