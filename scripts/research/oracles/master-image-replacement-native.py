#!/usr/bin/env python3
"""Independently inspect shared master replacement and four native slideshow samples."""
import json,posixpath,zipfile,xml.etree.ElementTree as ET
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from PIL import Image
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/".tmp/master-image-replacement-native-01"
NS={"p":"http://schemas.openxmlformats.org/presentationml/2006/main","a":"http://schemas.openxmlformats.org/drawingml/2006/main","r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships","rel":"http://schemas.openxmlformats.org/package/2006/relationships"}
def require(condition,message):
 if not condition:raise ValueError(message)
def digest(data):return sha256(data).hexdigest()
def inspect(package,pages,png):
 with Image.open(BytesIO(png)) as image:
  require(image.size==(4,2) and image.convert("RGBA").getpixel((0,0))==(255,0,0,255) and image.convert("RGBA").getpixel((3,0))[3]==0,"transparent fixture")
 with zipfile.ZipFile(BytesIO(package)) as z:
  def related(part,kind):
   relpath=posixpath.join(posixpath.dirname(part),"_rels",posixpath.basename(part)+".rels");rels=ET.fromstring(z.read(relpath))
   matches=[r for r in rels.findall("rel:Relationship",NS) if r.get("Type","").endswith("/"+kind)]
   require(len(matches)==1,"relationship cardinality "+part+" "+kind)
   return posixpath.normpath(posixpath.join(posixpath.dirname(part),matches[0].get("Target")))
  masters=[]
  for i in range(1,5):
   part=f"ppt/slides/slide{i}.xml";slide=ET.fromstring(z.read(part))
   require("".join(x.text or "" for x in slide.findall(".//a:t",NS))==f"MASTER IMAGE PAGE {i}","page identity")
   require(slide.get("showMasterSp") == ("0" if i==3 else "1"),"visibility")
   require(not slide.findall(".//p:pic",NS),"slide picture copy")
   masters.append(related(related(part,"slideLayout"),"slideMaster"))
  require(masters[0]==masters[1]==masters[2] and masters[3]!=masters[0],"shared owner graph")
  master=ET.fromstring(z.read(masters[0]));pictures=master.findall(".//p:pic",NS);require(len(pictures)==1,"master picture")
  pic=pictures[0];require(pic.find("p:nvPicPr/p:cNvPr",NS).get("name")=="logo","master picture identity")
  xfrm=pic.find("p:spPr/a:xfrm",NS);require(xfrm.find("a:off",NS).attrib=={"x":"1524000","y":"1270000"} and xfrm.find("a:ext",NS).attrib=={"cx":"2032000","cy":"3048000"},"frame")
  media=related(masters[0],"image");require(z.read(media)==png,"master image bytes")
  require(not ET.fromstring(z.read(masters[3])).findall(".//p:pic",NS),"other master picture")
 colors=[];titleHashes=[]
 for data in pages:
  with Image.open(BytesIO(data)) as image:
   image=image.convert("RGB");require(image.size==(1920,1080),"screenshot size")
   def classify(point):
    r,g,b=image.getpixel(point)
    if r>200 and g<80 and b<80:return "red"
    if r<80 and g<80 and b>200:return "blue"
    if r<80 and g>200 and b<80:return "green"
    if min(r,g,b)>240:return "white"
    return "other"
   colors.append([classify((300,400)),classify((510,400))]);titleHashes.append(digest(image.crop((80,40,550,110)).tobytes()))
 require(colors==[["red","blue"],["red","blue"],["white","white"],["green","green"]],"native sampled colors")
 require(len(set(titleHashes))==4,"distinct slide title crops")
 return {"masters":masters,"media":media,"sampledColors":colors,"distinctTitleCrops":len(set(titleHashes))}
def main():
 package=(BASE/"master-image-candidate.pptx").read_bytes();pages=[(BASE/f"native-page-{i}.png").read_bytes() for i in range(1,5)]
 png=(ROOT/"tests/fixtures/transparent-replacement.png").read_bytes();report=json.loads((BASE/"producer.json").read_text())
 require(digest(package)==report["pptxSha256"] and digest(png)==report["newSha256"] and report["warnings"]==[],"producer binding")
 result=inspect(package,pages,png);cases=[(package,[pages[0],pages[0],pages[2],pages[3]],png),(package,[pages[0],pages[1],pages[3],pages[2]],png),(package,pages,(ROOT/"tests/fixtures/exif-orientation6-normalized.png").read_bytes())]
 rejected=0
 for args in cases:
  try:inspect(*args)
  except Exception:rejected+=1
 require(rejected==3,"negative controls")
 paths=["scripts/prepare-master-image-replacement-native.mjs","packages/layout/src/image/plan-shared-visible-replacement.ts"]
 evidence={"scope":"PowerPoint 16.104 manual first-open slideshow on four named pages; sampled colors and title distinction, not an independent UI witness or native edit/save", "directory":str(BASE.relative_to(ROOT)),"pptxSha256":digest(package),"screenshotSha256":[digest(p) for p in pages],"producerSha256":digest((BASE/"producer.json").read_bytes()),"transparentFixtureSha256":digest(png),"scripts":{p:digest((ROOT/p).read_bytes()) for p in paths},"oracleSha256":digest(Path(__file__).read_bytes()),"observation":result,"negativeCasesRejected":rejected}
 (ROOT/"docs/testing/master-image-replacement-native-evidence.json").write_text(json.dumps(evidence,indent=2)+"\n");print(json.dumps(result))
if __name__=="__main__":main()
