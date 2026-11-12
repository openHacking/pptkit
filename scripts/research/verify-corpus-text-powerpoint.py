#!/usr/bin/env python3
"""Independently inspect a bounded PowerPoint text-corpus observation."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"a": A, "p": P}


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def inspect(path, edited):
    with zipfile.ZipFile(path) as package:
        slide = ET.fromstring(package.read("ppt/slides/slide1.xml"))
        second = ET.fromstring(package.read("ppt/slides/slide2.xml"))
        notes = ET.fromstring(package.read("ppt/notesSlides/notesSlide2.xml"))
        rels = ET.fromstring(package.read("ppt/slides/_rels/slide1.xml.rels"))
    text = lambda node: "".join(t.text or "" for t in node.findall(".//a:t", NS))
    paragraphs = [p for p in slide.findall(".//a:p", NS) if p.findall(".//a:t", NS)]
    expected = (["Edited & <editable> text", "Revised paragraph"] if edited else
                ["Rich & <editable> text", "Numbered paragraph"]) + ["Bulleted paragraph", "Go to notes"]
    if [text(p) for p in paragraphs] != expected:
        raise ValueError("paragraphs differ")
    num = paragraphs[1].find("a:pPr/a:buAutoNum", NS)
    bullet = paragraphs[2].find("a:pPr/a:buChar", NS)
    if num is None or num.attrib.get("type") != "arabicPeriod" or num.attrib.get("startAt") != "3" or bullet is None or bullet.attrib.get("char") != "*":
        raise ValueError("list properties differ")
    if second.get("show") != "0" or "Speaker & reviewer notes" not in text(notes):
        raise ValueError("hidden slide or notes differ")
    links = [node for p in (paragraphs[0], paragraphs[3]) for node in p.findall(".//a:hlinkClick", NS)]
    if len(links) != 2 or links[1].get("action") != "ppaction://hlinksldjump":
        raise ValueError("link actions differ")
    relationships = {el.get("Id"): el for el in rels.findall(f"{{{PKG}}}Relationship")}
    external, internal = [relationships.get(link.get(f"{{{R}}}id")) for link in links]
    if external is None or external.get("Target") != "https://example.com/?a=1&b=2" or external.get("TargetMode") != "External":
        raise ValueError("external link differs")
    if internal is None or internal.get("Target") != "slide2.xml" or not internal.get("Type", "").endswith("/slide"):
        raise ValueError("slide link differs")


def verify(local=False):
    record = json.loads((ROOT / "docs/testing/corpus-text-powerpoint-evidence.json").read_text())
    output = ROOT / record["output"]["path"]
    if digest(output) != record["output"]["sha256"]:
        raise ValueError("output hash mismatch")
    inspect(output, True)
    if local:
        source = ROOT / record["input"]["path"]
        manifest = json.loads((ROOT / record["input"]["manifest"]).read_text())
        entry = next(item for item in manifest["fixtures"] if item["id"] == "text")
        if digest(source) != record["input"]["sha256"] or entry["sha256"] != record["input"]["sha256"]:
            raise ValueError("source or manifest hash mismatch")
        inspect(source, False)
        for name, expected in record["captures"].items():
            if digest(ROOT / record["artifactDirectory"] / name) != expected:
                raise ValueError(f"capture hash mismatch: {name}")


if __name__ == "__main__":
    if sys.argv[1:] not in ([], ["--local"]):
        raise SystemExit("usage: verify-corpus-text-powerpoint.py [--local]")
    verify(bool(sys.argv[1:]))
    print("verified native text package" + (" and local evidence" if sys.argv[1:] else ""))
