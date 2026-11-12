#!/usr/bin/env python3
"""Verify the sampled native image-crop save and its shared media."""
import hashlib
import json
from pathlib import Path
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def digest(data):
    return hashlib.sha256(data).hexdigest()


def inspect(path, edited, media_hash):
    with zipfile.ZipFile(path) as archive:
        slide = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        rels = ET.fromstring(archive.read("ppt/slides/_rels/slide1.xml.rels"))
        relationships = {x.get("Id"): x for x in rels.findall(f"{{{REL}}}Relationship")}
        pictures = {}
        for pic in slide.findall(".//p:pic", NS):
            name = pic.find("p:nvPicPr/p:cNvPr", NS).get("name")
            if name in pictures:
                raise ValueError("duplicate picture name")
            rid = pic.find("p:blipFill/a:blip", NS).get(f"{{{R}}}embed")
            rel = relationships.get(rid)
            if rel is None or not rel.get("Type", "").endswith("/image") or rel.get("TargetMode") == "External":
                raise ValueError("picture relationship differs")
            part = posixpath.normpath(posixpath.join("ppt/slides", rel.get("Target", "")))
            if digest(archive.read(part)) != media_hash:
                raise ValueError("picture media differs")
            crop = pic.find("p:blipFill/a:srcRect", NS)
            xfrm = pic.find("p:spPr/a:xfrm", NS)
            off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
            pictures[name] = ({key: int(crop.get(key, "0")) for key in ("l", "t", "r", "b")} if crop is not None else None,
                              [int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))])
    if set(pictures) != {"image-stretch", "image-contain", "image-cover", "image-crop"}:
        raise ValueError("picture inventory differs")
    expected_crop = {"l": 18212, "t": 0, "r": 0, "b": 0} if edited else None
    expected_frame = [924338, 1016000, 1869661, 1397000] if edited else [508000, 1016000, 2286000, 1397000]
    if pictures["image-stretch"] != (expected_crop, expected_frame):
        raise ValueError("edited crop or geometry differs")
    if pictures["image-contain"] != (None, [3302000, 1143000, 2286000, 1143000]):
        raise ValueError("contain control differs")
    if pictures["image-cover"] != ({"l": 9091, "t": 0, "r": 9091, "b": 0}, [6096000, 1016000, 2286000, 1397000]):
        raise ValueError("cover control differs")
    if pictures["image-crop"] != ({"l": 10000, "t": 15000, "r": 20000, "b": 5000}, [8890000, 1016000, 2286000, 1397000]):
        raise ValueError("crop control differs")


def verify(local=False):
    data = json.loads((ROOT / "docs/testing/corpus-images-powerpoint-evidence.json").read_text())
    output = ROOT / data["output"]["path"]
    if digest(output.read_bytes()) != data["output"]["sha256"]:
        raise ValueError("output hash mismatch")
    inspect(output, True, data["mediaSha256"])
    if local:
        source = ROOT / data["input"]["path"]
        manifest = json.loads((ROOT / data["input"]["manifest"]).read_text())
        item = next(x for x in manifest["fixtures"] if x["id"] == "images")
        if digest(source.read_bytes()) != data["input"]["sha256"] or item["sha256"] != data["input"]["sha256"]:
            raise ValueError("source or manifest hash mismatch")
        inspect(source, False, data["mediaSha256"])
        for name, expected in data["captures"].items():
            if digest((ROOT / data["artifactDirectory"] / name).read_bytes()) != expected:
                raise ValueError(f"capture hash mismatch: {name}")


if __name__ == "__main__":
    if sys.argv[1:] not in ([], ["--local"]):
        raise SystemExit("usage: verify-corpus-images-powerpoint.py [--local]")
    verify(bool(sys.argv[1:]))
    print("verified image package" + (" and local evidence" if sys.argv[1:] else ""))
