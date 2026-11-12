#!/usr/bin/env python3
"""Check a named fresh projection and PowerPoint Save As/reopen sample."""
import importlib.util,json,posixpath,zipfile,xml.etree.ElementTree as ET
from io import BytesIO
from pathlib import Path
from hashlib import sha256
ROOT=Path(__file__).resolve().parents[3];BASE=ROOT/".tmp/master-artwork-projection-native-01"
path=ROOT/"scripts/research/oracles/master-image-native-edit.py";spec=importlib.util.spec_from_file_location("previous_oracle",path);oracle=importlib.util.module_from_spec(spec);spec.loader.exec_module(oracle)
digest=lambda b:sha256(b).hexdigest();require=oracle.require;ns=oracle.NS
def inspect(package,pages,png):
 result=oracle.inspect(package,pages,png)
 with zipfile.ZipFile(BytesIO(package)) as z:
  master=ET.fromstring(z.read(result["masters"][0]));tree=master.find("p:cSld/p:spTree",ns)
  shapes=[n for n in tree if n.tag in (f'{{{ns["p"]}}}sp',f'{{{ns["p"]}}}pic')]
  require([x.find("p:nvSpPr/p:cNvPr",ns).get("name") if x.tag.endswith("}sp") else x.find("p:nvPicPr/p:cNvPr",ns).get("name") for x in shapes]==["blue","logo"],"master drawing order")
  require(shapes[0].find("p:spPr/a:solidFill/a:srgbClr",ns).get("val").upper()=="0000FF","blue fill")
  other=ET.fromstring(z.read(result["masters"][3]));green=next(x for x in other.findall(".//p:sp",ns) if x.find("p:nvSpPr/p:cNvPr",ns).get("name")=="green")
  require(green.find("p:spPr/a:solidFill/a:srgbClr",ns).get("val").upper()=="00FF00","green fill")
 return result
def main():
 first=(BASE/"projection.pptx").read_bytes();saved=(BASE/"projection-powerpoint-saved.pptx").read_bytes();png=(ROOT/"tests/fixtures/transparent-replacement.png").read_bytes()
 screenshots={stage:[(BASE/f"{stage}-page-{i}.png").read_bytes() for i in range(1,5)] for stage in ("first","reopen")}
 observations={"first":inspect(first,screenshots["first"],png),"reopen":inspect(saved,screenshots["reopen"],png)}
 require(digest(first)!=digest(saved),"PowerPoint output is not a distinct serialized copy")
 old=(ROOT/".tmp/master-image-replacement-native-01/master-image-candidate.pptx").read_bytes()
 controls=[(old,screenshots["first"],png),(saved,[screenshots["reopen"][0],screenshots["reopen"][0],*screenshots["reopen"][2:]],png),(saved,screenshots["reopen"],b"wrong bytes")]
 rejected=0
 for args in controls:
  try:inspect(*args)
  except Exception:rejected+=1
 require(rejected==len(controls),"negative controls")
 evidence={"scope":"PowerPoint 16.104 opens a fresh partial-projection export, displays four pages, Save As, closes and reopens the saved copy; screenshot/title samples are not independent UI witnesses; no native edit or source-preserving save", "directory":str(BASE.relative_to(ROOT)),"sourceSha256":digest((ROOT/".tmp/master-image-replacement-native-01/master-image-native-edit.pptx").read_bytes()),"projectedSha256":digest(first),"powerpointSavedSha256":digest(saved),"screenshots":{k:[digest(b) for b in v] for k,v in screenshots.items()},"producerSha256":digest((ROOT/"scripts/prepare-master-artwork-projection-native.mjs").read_bytes()),"priorOracleSha256":digest(path.read_bytes()),"oracleSha256":digest(Path(__file__).read_bytes()),"observation":observations,"negativeCasesRejected":rejected}
 (ROOT/"docs/testing/master-artwork-projection-native-evidence.json").write_text(json.dumps(evidence,indent=2)+"\n");print(json.dumps(observations))
if __name__=="__main__":main()
