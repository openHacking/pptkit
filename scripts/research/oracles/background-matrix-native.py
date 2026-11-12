#!/usr/bin/env python3
"""Independently inspect a named bgRef projection and native slideshow sample."""
import json,posixpath,zipfile,xml.etree.ElementTree as ET
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from PIL import Image,ImageChops
ROOT=Path(__file__).resolve().parents[3]; BASE=ROOT/".tmp/background-matrix-native-01"
N={"p":"http://schemas.openxmlformats.org/presentationml/2006/main","a":"http://schemas.openxmlformats.org/drawingml/2006/main","r":"http://schemas.openxmlformats.org/package/2006/relationships"}
digest=lambda b:sha256(b).hexdigest()
def require(value,message):
 if not value:raise AssertionError(message)
def inspect(data,source):
 with zipfile.ZipFile(BytesIO(data)) as z:
  def rel(part,role):
   path=posixpath.join(posixpath.dirname(part),"_rels",posixpath.basename(part)+".rels");items=[x for x in ET.fromstring(z.read(path)).findall("r:Relationship",N) if x.get("Type","").endswith("/"+role)]
   require(len(items)==1 and items[0].get("TargetMode","Internal")=="Internal","relationship");return posixpath.normpath(posixpath.join(posixpath.dirname(part),items[0].get("Target")))
  colors=["A00000","008000","0000C0"];owners=[]
  for index in range(1,4):
   slide=f"ppt/slides/slide{index}.xml";s=ET.fromstring(z.read(slide));require("".join(t.text or "" for t in s.findall(".//a:t",N))==f"BGREF PAGE {index}","page text")
   layout=rel(slide,"slideLayout");master=rel(layout,"slideMaster");parts=[slide,layout,master];backgrounds=[ET.fromstring(z.read(p)).find("p:cSld/p:bg",N) for p in parts]
   selected=next((i for i,b in enumerate(backgrounds) if b is not None),None);require(selected==[2,1,0][index-1],"background ownership");bg=backgrounds[selected]
   if source:
    ref=bg.find("p:bgRef",N);require(ref is not None and ref.get("idx")=="1001","background matrix index");c=ref.find("a:srgbClr",N);require(c is not None and c.get("val")==colors[index-1],"placeholder color")
    theme=ET.fromstring(z.read(rel(master,"theme")));matrix=theme.find("a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:solidFill/a:schemeClr",N);require(matrix is not None and matrix.get("val")=="phClr","matrix placeholder")
   else:
    fill=bg.find("p:bgPr/a:solidFill/a:srgbClr",N);require(fill is not None and fill.get("val")==colors[index-1],"materialized background")
   owners.append({"slide":slide,"layout":layout,"master":master,"owner":parts[selected],"color":colors[index-1]})
  require(owners[0]["master"]==owners[1]["master"]==owners[2]["master"],"shared master");require(owners[0]["layout"]!=owners[1]["layout"]==owners[2]["layout"],"layout topology")
  return owners
def screenshots(stages):
 expected=[(153,0,3),(0,115,24),(0,27,180)];samples={};titles=[]
 for stage,pages in stages.items():
  samples[stage]=[]
  for i,data in enumerate(pages):
   image=Image.open(BytesIO(data)).convert("RGB");require(image.size==(1920,1080),"screenshot size");pixel=image.getpixel((900,500));require(all(abs(a-b)<=2 for a,b in zip(pixel,expected[i])),"background sample");samples[stage].append(pixel)
   title=image.crop((70,50,600,160));baseline=Image.open(BytesIO(stages["source"][i])).convert("RGB").crop((70,50,600,160));require(ImageChops.difference(title,baseline).getbbox() is None,"title or background crop")
   if stage=="source":titles.append(digest(title.tobytes()))
 require(len(set(titles))==3,"distinct title crops");return samples
def main():
 names={"source":"source-bgref.pptx","projected":"projected-bgref.pptx","reopen":"projected-powerpoint-saved.pptx"};files={k:(BASE/v).read_bytes() for k,v in names.items()};pages={k:[(BASE/f"{k}-page-{i}.png").read_bytes() for i in range(1,4)] for k in names}
 observations={k:inspect(v,k=="source") for k,v in files.items()};samples=screenshots(pages);require(digest(files["projected"])!=digest(files["reopen"]),"PowerPoint serialized copy")
 controls=[lambda:inspect(files["projected"],True),lambda:inspect(files["source"],False),lambda:screenshots({**pages,"reopen":[pages["reopen"][0],pages["reopen"][0],pages["reopen"][2]]})];rejected=0
 for control in controls:
  try:control()
  except Exception:rejected+=1
 require(rejected==len(controls),"negative controls")
 evidence={"scope":"PowerPoint 16.104 slideshow of authored bgRef source and partial projection, then Save As/close/reopen of projection; screenshot samples do not independently establish UI provenance; no native object edit or preservation save", "directory":str(BASE.relative_to(ROOT)),"files":{k:digest(v) for k,v in files.items()},"screenshots":{k:[digest(x) for x in v] for k,v in pages.items()},"producerSha256":digest((ROOT/"scripts/prepare-background-matrix-native.mjs").read_bytes()),"oracleSha256":digest(Path(__file__).read_bytes()),"owners":observations,"sampledRgb":samples,"negativeCasesRejected":rejected}
 (ROOT/"docs/testing/background-matrix-native-evidence.json").write_text(json.dumps(evidence,indent=2)+"\n");print(json.dumps({"owners":observations,"samples":samples}))
if __name__=="__main__":main()
