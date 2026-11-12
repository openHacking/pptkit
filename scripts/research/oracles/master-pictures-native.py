#!/usr/bin/env python3
"""Compare the public inventory against raw OOXML in a named native-produced deck."""
import copy,json,posixpath,zipfile,xml.etree.ElementTree as ET
from io import BytesIO
from hashlib import sha256
from pathlib import Path
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/".tmp/master-image-replacement-native-01"
NS={"p":"http://schemas.openxmlformats.org/presentationml/2006/main","a":"http://schemas.openxmlformats.org/drawingml/2006/main","r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships","rel":"http://schemas.openxmlformats.org/package/2006/relationships"}
def check(condition,message):
 if not condition:raise ValueError(message)
def digest(data):return sha256(data).hexdigest()
def inspect(report,package):
 check(report["inputSha256"]==digest(package),"input hash");inventory=report["inventory"]
 check(inventory["coverage"]=="effective-listed-masters-only" and inventory["mapping"]=="not-performed","scope")
 owners=[m for m in inventory["masters"] if any(p.get("name")=="logo" for p in m["pictures"])];check(len(owners)==1,"owner count");owner=owners[0];item=next(p for p in owner["pictures"] if p.get("name")=="logo")
 with zipfile.ZipFile(BytesIO(package)) as z:
  part=owner["part"].lstrip("/");root=ET.fromstring(z.read(part));nodes=list(root.iter());node=nodes[item["source"]["elementIndex"]-1]
  check(node.tag=="{%s}pic"%NS["p"] and item["source"]["part"]==owner["part"],"raw index")
  cnv=node.find("p:nvPicPr/p:cNvPr",NS);check(cnv.get("name")==item["name"] and cnv.get("id")==item["sourceObjectId"],"nonvisual identity")
  rect=node.find("p:blipFill/a:srcRect",NS);check(item["crop"]=={"left":rect.get("l")} and item["crop"]["left"]=="20870","crop")
  xfrm=node.find("p:spPr/a:xfrm",NS);off=xfrm.find("a:off",NS);ext=xfrm.find("a:ext",NS)
  check(item["transform"]=={"x":off.get("x"),"y":off.get("y"),"cx":ext.get("cx"),"cy":ext.get("cy")},"transform")
  rid=node.find("p:blipFill/a:blip",NS).get("{%s}embed"%NS["r"]);relpath=posixpath.join(posixpath.dirname(part),"_rels",posixpath.basename(part)+".rels")
  rels=ET.fromstring(z.read(relpath));rel=next(r for r in rels.findall("rel:Relationship",NS) if r.get("Id")==rid)
  target="/"+posixpath.normpath(posixpath.join(posixpath.dirname(part),rel.get("Target")))
  check(item["embed"]=={"id":rid,"status":"internal-image","targetPart":target},"image relationship")
  check(owner["slideIds"]==[256,257,258] and item["groupElementIndices"]==[],"owner membership")
 return {"part":owner["part"],"rawElementIndex":item["source"]["elementIndex"],"crop":item["crop"],"target":target,"slideIds":owner["slideIds"]}
def main():
 package=(BASE/"master-image-native-edit.pptx").read_bytes();raw=(BASE/"master-picture-inventory.json").read_bytes();report=json.loads(raw);observation=inspect(report,package)
 rejected=0
 for field,value in [("sourceObjectId","999"),("crop",{"left":"0"}),("groupElementIndices",[1])]:
  changed=copy.deepcopy(report);next(p for m in changed["inventory"]["masters"] for p in m["pictures"] if p.get("name")=="logo")[field]=value
  try:inspect(changed,package)
  except Exception:rejected+=1
 check(rejected==3,"negative controls")
 scripts=["packages/pptx-parser/src/master-pictures/inventory.ts","packages/pptx-parser/src/types/master-pictures.ts","packages/pptx-parser/src/index.ts","scripts/inspect-master-pictures-native.mjs"]
 evidence={"scope":"source inventory cross-check on one PowerPoint 16.104 native-edited file; no semantic projection, rendering or source-edit permission", "directory":str(BASE.relative_to(ROOT)),"nativeEvidenceSha256":digest((ROOT/"docs/testing/master-image-native-edit-evidence.json").read_bytes()),"inventoryReportSha256":digest(raw),"inputSha256":digest(package),"oracleSha256":digest(Path(__file__).read_bytes()),"scripts":{x:digest((ROOT/x).read_bytes()) for x in scripts},"observation":observation,"negativeCasesRejected":rejected}
 (ROOT/"docs/testing/master-pictures-native-evidence.json").write_text(json.dumps(evidence,indent=2)+"\n");print(json.dumps(observation))
if __name__=="__main__":main()
