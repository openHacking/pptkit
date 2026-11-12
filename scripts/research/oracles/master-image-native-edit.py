#!/usr/bin/env python3
"""Verify one named PowerPoint master picture crop edit, save and reopen."""
import importlib.util,json,posixpath,zipfile,xml.etree.ElementTree as ET
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from PIL import Image
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/".tmp/master-image-replacement-native-01"
path=ROOT/"scripts/research/oracles/master-image-replacement-native.py";spec=importlib.util.spec_from_file_location("reference",path);reference=importlib.util.module_from_spec(spec);spec.loader.exec_module(reference)
NS=reference.NS;require=reference.require;digest=reference.digest

def inspect(package,pages,png):
 with zipfile.ZipFile(BytesIO(package)) as z:
  def related(part,kind):
   relpath=posixpath.join(posixpath.dirname(part),"_rels",posixpath.basename(part)+".rels");root=ET.fromstring(z.read(relpath));matches=[r for r in root.findall("rel:Relationship",NS) if r.get("Type","").endswith("/"+kind)]
   require(len(matches)==1,"relationship cardinality");return posixpath.normpath(posixpath.join(posixpath.dirname(part),matches[0].get("Target")))
  masters=[]
  for i in range(1,5):
   part=f"ppt/slides/slide{i}.xml";slide=ET.fromstring(z.read(part));require("".join(t.text or "" for t in slide.findall(".//a:t",NS))==f"MASTER IMAGE PAGE {i}","slide text")
   require(slide.get("showMasterSp","1")==("0" if i==3 else "1"),"slide visibility");require(not slide.findall(".//p:pic",NS),"no per-slide picture")
   masters.append(related(related(part,"slideLayout"),"slideMaster"))
  require(masters[0]==masters[1]==masters[2] and masters[3]!=masters[0],"master graph")
  master=ET.fromstring(z.read(masters[0]));pics=master.findall(".//p:pic",NS);require(len(pics)==1,"picture count");pic=pics[0]
  require(pic.find("p:nvPicPr/p:cNvPr",NS).get("name")=="logo","picture identity")
  crop=pic.find("p:blipFill/a:srcRect",NS);require(crop is not None and {k:int(crop.get(k,"0")) for k in "ltrb"}=={"l":20870,"t":0,"r":0,"b":0},"edited crop")
  xfrm=pic.find("p:spPr/a:xfrm",NS);require(xfrm.find("a:off",NS).attrib=={"x":"1948070","y":"1270000"} and xfrm.find("a:ext",NS).attrib=={"cx":"1607930","cy":"3048000"},"edited frame")
  media=related(masters[0],"image");require(z.read(media)==png,"embedded image bytes")
  other=ET.fromstring(z.read(masters[3]));require(not other.findall(".//p:pic",NS),"other master")
  require(any(s.find("p:nvSpPr/p:cNvPr",NS).get("name")=="green" for s in other.findall(".//p:sp",NS)),"other master green shape")
 colors=[];titleHashes=[]
 for index,data in enumerate(pages,1):
  with Image.open(BytesIO(data)) as img, Image.open(BASE/f"native-page-{index}.png") as prior:
   image=img.convert("RGB");prior=prior.convert("RGB");require(image.size==(1920,1080),"screenshot size")
   require(image.crop((80,40,550,110)).tobytes()==prior.crop((80,40,550,110)).tobytes(),"page title crop differs from first-open sample")
   def color(point):
    r,g,b=image.getpixel(point)
    if r>200 and g<80 and b<80:return "red"
    if r<80 and g<80 and b>200:return "blue"
    if r<80 and g>200 and b<80:return "green"
    if min(r,g,b)>240:return "white"
    return "other"
   colors.append([color((270,400)),color((340,400)),color((510,400))]);titleHashes.append(digest(image.crop((80,40,550,110)).tobytes()))
 require(colors==[["blue","red","blue"],["blue","red","blue"],["white","white","white"],["green","green","green"]],"edited page samples")
 require(len(set(titleHashes))==4,"distinct titles")
 return {"masters":masters,"media":media,"crop":{"left":20870},"sampledColors":colors}

def main():
 original=(BASE/"master-image-candidate.pptx").read_bytes();package=(BASE/"master-image-native-edit.pptx").read_bytes();png=(ROOT/"tests/fixtures/transparent-replacement.png").read_bytes()
 before=[(BASE/f"native-page-{i}.png").read_bytes() for i in range(1,5)];pages=[(BASE/f"edit-page-{i}.png").read_bytes() for i in range(1,5)]
 reference.inspect(original,before,png);require(digest(original)!=digest(package),"native file was rewritten");observation=inspect(package,pages,png)
 parserBytes=(BASE/"native-edit-parser.json").read_bytes();parser=json.loads(parserBytes)
 require(parser["inputSha256"]==digest(package) and parser["projectedMasterCount"]==0 and parser["noOpByteIdentical"] is True,"parser report")
 cases=[(original,pages,png),(package,before,png),(package,[pages[0],pages[0],pages[2],pages[3]],png)]
 rejected=0
 for args in cases:
  try:inspect(*args)
  except Exception:rejected+=1
 require(rejected==3,"negative controls")
 evidence={"scope":"PowerPoint 16.104 manual master image crop-handle edit, save, close, reopen and four-page slideshow; screenshot samples and title regions are not independent UI witnesses", "directory":str(BASE.relative_to(ROOT)),"previousEvidenceSha256":digest((ROOT/"docs/testing/master-image-replacement-native-evidence.json").read_bytes()),"editedPptxSha256":digest(package),"screenshotSha256":[digest(x) for x in pages],"referenceOracleSha256":digest(path.read_bytes()),"oracleSha256":digest(Path(__file__).read_bytes()),"parserReportSha256":digest(parserBytes),"parserScriptSha256":digest((ROOT/"scripts/inspect-master-image-native-edit.mjs").read_bytes()),"observation":observation,"negativeCasesRejected":rejected}
 (ROOT/"docs/testing/master-image-native-edit-evidence.json").write_text(json.dumps(evidence,indent=2)+"\n");print(json.dumps(observation))
if __name__=="__main__":main()
