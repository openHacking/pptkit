#!/usr/bin/env python3
"""Verify the saved corpus theme/title observation and optional local captures."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def inspect(path, title):
    with zipfile.ZipFile(path) as archive:
        slide = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        hidden = ET.fromstring(archive.read("ppt/slides/slide2.xml"))
        layout = ET.fromstring(archive.read("ppt/slideLayouts/slideLayout2.xml"))
        ET.fromstring(archive.read("ppt/slideMasters/slideMaster1.xml"))
        ET.fromstring(archive.read("ppt/theme/theme1.xml"))
        rels = [ET.fromstring(archive.read(name)) for name in (
            "ppt/slides/_rels/slide1.xml.rels", "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
            "ppt/slideMasters/_rels/slideMaster1.xml.rels")]
    shapes = [shape for shape in slide.findall(".//p:sp", NS)
              if (prop := shape.find("p:nvSpPr/p:cNvPr", NS)) is not None and prop.get("name") == "bound-title"]
    if len(shapes) != 1 or "".join(node.text or "" for node in shapes[0].findall(".//a:t", NS)) != title:
        raise ValueError("title differs")
    ph = shapes[0].find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is None or ph.get("type") != "title" or ph.get("idx") != "1" or hidden.get("show") != "0":
        raise ValueError("placeholder or visibility differs")
    if not any(x.get("type") == "title" and x.get("idx") == "1" for x in layout.findall(".//p:ph", NS)):
        raise ValueError("layout title placeholder missing")
    for root, kind, target in zip(rels, ("slideLayout", "slideMaster", "theme"),
                                  ("../slideLayouts/slideLayout2.xml", "../slideMasters/slideMaster1.xml", "../theme/theme1.xml")):
        matches = [x for x in root.findall(f"{{{REL}}}Relationship")
                   if x.get("Type", "").endswith("/" + kind) and x.get("Target") == target]
        if len(matches) != 1:
            raise ValueError(f"{kind} relationship differs")


def verify(local=False):
    data = json.loads((ROOT / "docs/testing/corpus-theme-powerpoint-evidence.json").read_text())
    output = ROOT / data["output"]["path"]
    if digest(output) != data["output"]["sha256"]:
        raise ValueError("output hash mismatch")
    inspect(output, "Edited title")
    if local:
        source = ROOT / data["input"]["path"]
        manifest = json.loads((ROOT / data["input"]["manifest"]).read_text())
        item = next(x for x in manifest["fixtures"] if x["id"] == "theme-layout")
        if digest(source) != data["input"]["sha256"] or item["sha256"] != data["input"]["sha256"]:
            raise ValueError("source or manifest hash mismatch")
        inspect(source, "Inherited title")
        for name, expected in data["captures"].items():
            if digest(ROOT / data["artifactDirectory"] / name) != expected:
                raise ValueError(f"capture hash mismatch: {name}")


if __name__ == "__main__":
    if sys.argv[1:] not in ([], ["--local"]):
        raise SystemExit("usage: verify-corpus-theme-powerpoint.py [--local]")
    verify(bool(sys.argv[1:]))
    print("verified theme package" + (" and local evidence" if sys.argv[1:] else ""))
