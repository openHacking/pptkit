#!/usr/bin/env python3
"""Check the retained PowerPoint group-movement observation independently."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def inspect(path, offset):
    with zipfile.ZipFile(path) as archive:
        slide = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
    groups = slide.findall(".//p:grpSp", NS)
    if len(groups) != 2:
        raise ValueError("group count differs")
    outer, inner = groups
    def group_name(group):
        return group.find("p:nvGrpSpPr/p:cNvPr", NS).get("name")
    if group_name(outer) != "outer" or group_name(inner) != "inner" or inner not in list(outer):
        raise ValueError("group hierarchy differs")
    ox = outer.find("p:grpSpPr/a:xfrm", NS)
    ix = inner.find("p:grpSpPr/a:xfrm", NS)
    position = ox.find("a:off", NS)
    if [int(position.get("x")), int(position.get("y"))] != offset or ox.get("rot") != "600000" or ix.get("flipH") != "1":
        raise ValueError("group transform differs")
    if [(e.tag.rsplit("}", 1)[-1], e.attrib) for e in ix] != [
        ("off", {"x": "127000", "y": "127000"}), ("ext", {"cx": "1778000", "cy": "762000"}),
        ("chOff", {"x": "0", "y": "0"}), ("chExt", {"cx": "3556000", "cy": "1524000"})]:
        raise ValueError("inner coordinates differ")
    names = [node.get("name") for node in inner.findall("p:sp/p:nvSpPr/p:cNvPr", NS)]
    if names != ["nested-shape", "nested-text"] or "Nested" not in [node.text for node in inner.findall(".//a:t", NS)]:
        raise ValueError("group children differ")
    shapes = [shape for shape in slide.findall(".//p:sp", NS)
              if (name := shape.find("p:nvSpPr/p:cNvPr", NS)) is not None and name.get("name") == "anchored"]
    if len(shapes) != 1 or shapes[0].find("p:spPr/a:custGeom", NS) is None:
        raise ValueError("ordinary path missing")


def verify(local=False):
    record = json.loads((ROOT / "docs/testing/corpus-geometry-powerpoint-evidence.json").read_text())
    output = ROOT / record["output"]["path"]
    if digest(output) != record["output"]["sha256"]:
        raise ValueError("output hash mismatch")
    inspect(output, record["outerOffsetAfterEMU"])
    if local:
        source = ROOT / record["input"]["path"]
        manifest = json.loads((ROOT / record["input"]["manifest"]).read_text())
        entry = next(item for item in manifest["fixtures"] if item["id"] == "geometry")
        if digest(source) != record["input"]["sha256"] or entry["sha256"] != record["input"]["sha256"]:
            raise ValueError("source or manifest hash mismatch")
        inspect(source, record["outerOffsetBeforeEMU"])
        for name, expected in record["captures"].items():
            if digest(ROOT / record["artifactDirectory"] / name) != expected:
                raise ValueError(f"capture hash mismatch: {name}")


if __name__ == "__main__":
    if sys.argv[1:] not in ([], ["--local"]):
        raise SystemExit("usage: verify-corpus-geometry-powerpoint.py [--local]")
    verify(bool(sys.argv[1:]))
    print("verified geometry package" + (" and local evidence" if sys.argv[1:] else ""))
