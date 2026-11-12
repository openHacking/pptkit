#!/usr/bin/env python3
"""Check the missing-asset warning and the resulting native blank-page sample."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def digest(data):
    return hashlib.sha256(data).hexdigest()


def inspect(path):
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("ppt/slides/slide1.xml"))
        rels = ET.fromstring(package.read("ppt/slides/_rels/slide1.xml.rels"))
        if root.findall(f".//{{{P}}}pic") or any(name.startswith("ppt/media/") for name in package.namelist()):
            raise ValueError("unexpected picture or media")
        if any(x.get("Type", "").endswith("/image") for x in rels.findall(f"{{{REL}}}Relationship")):
            raise ValueError("unexpected image relationship")
        if any(root.findall(f".//{{{P}}}{tag}") for tag in ("sp", "grpSp", "graphicFrame", "cxnSp")):
            raise ValueError("unexpected slide content")


def verify(local=False):
    data = json.loads((ROOT / "docs/testing/corpus-missing-powerpoint-evidence.json").read_text())
    output = ROOT / data["output"]["path"]
    if digest(output.read_bytes()) != data["output"]["sha256"]:
        raise ValueError("output hash mismatch")
    inspect(output)
    if local:
        source = ROOT / data["input"]["path"]
        manifest = json.loads((ROOT / data["input"]["manifest"]).read_text())
        item = next(x for x in manifest["fixtures"] if x["id"] == "missing-asset")
        if digest(source.read_bytes()) != data["input"]["sha256"] or item["sha256"] != data["input"]["sha256"]:
            raise ValueError("source or manifest hash mismatch")
        if [x["code"] for x in item["exportWarnings"]] != ["asset-read-failed", "image-omitted"] or [x["code"] for x in item["previewWarnings"]] != ["asset-path-unsupported"]:
            raise ValueError("expected warnings missing")
        inspect(source)
        for name, expected in data["captures"].items():
            if digest((ROOT / data["artifactDirectory"] / name).read_bytes()) != expected:
                raise ValueError(f"capture hash mismatch: {name}")


if __name__ == "__main__":
    if sys.argv[1:] not in ([], ["--local"]):
        raise SystemExit("usage: verify-corpus-missing-powerpoint.py [--local]")
    verify(bool(sys.argv[1:]))
    print("verified missing-asset package" + (" and local evidence" if sys.argv[1:] else ""))
