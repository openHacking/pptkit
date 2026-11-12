#!/usr/bin/env python3
"""Check native chart title, caches and referenced embedded worksheet values."""
import hashlib
import io
import json
from pathlib import Path
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def digest(data):
    return hashlib.sha256(data).hexdigest()


def inspect(path, edited):
    with zipfile.ZipFile(path) as archive:
        targets = set()
        for i, title in enumerate((["Edited bar", "line", "pie"] if edited else ["bar", "line", "pie"]), 1):
            chart = ET.fromstring(archive.read(f"ppt/charts/chart{i}.xml"))
            text = "".join(x.text or "" for x in chart.findall(".//c:title//a:t", NS))
            points = chart.findall(".//c:val/c:numRef/c:numCache/c:pt", NS)
            if text != title or [(p.get("idx"), p.find("c:v", NS).text) for p in points] != [("0", "2"), ("1", "4"), ("2", "3")]:
                raise ValueError("chart title or cache differs")
            external = chart.find("c:externalData", NS)
            if external is None:
                raise ValueError("chart has no workbook reference")
            rid = external.get(f"{{{R}}}id")
            rels = ET.fromstring(archive.read(f"ppt/charts/_rels/chart{i}.xml.rels"))
            matches = [x for x in rels.findall(f"{{{REL}}}Relationship") if x.get("Id") == rid]
            if len(matches) != 1 or not matches[0].get("Type", "").endswith("/package") or matches[0].get("TargetMode") == "External":
                raise ValueError("workbook relationship differs")
            target = posixpath.normpath(posixpath.join("ppt/charts", matches[0].get("Target", "")))
            targets.add(target)
            with zipfile.ZipFile(io.BytesIO(archive.read(target))) as workbook:
                sheet = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
                values = {x.get("r"): x.find("s:v", NS).text for x in sheet.findall(".//s:c", NS) if x.find("s:v", NS) is not None}
            if {key: values.get(key) for key in ("B2", "B3", "B4")} != {"B2": "2", "B3": "4", "B4": "3"}:
                raise ValueError("workbook values differ")
        if len(targets) != 3:
            raise ValueError("workbook identities collapsed")


def verify(local=False):
    data = json.loads((ROOT / "docs/testing/corpus-charts-powerpoint-evidence.json").read_text())
    output = ROOT / data["output"]["path"]
    if digest(output.read_bytes()) != data["output"]["sha256"]:
        raise ValueError("output hash mismatch")
    inspect(output, True)
    if local:
        source = ROOT / data["input"]["path"]
        manifest = json.loads((ROOT / data["input"]["manifest"]).read_text())
        item = next(x for x in manifest["fixtures"] if x["id"] == "charts")
        if digest(source.read_bytes()) != data["input"]["sha256"] or item["sha256"] != data["input"]["sha256"]:
            raise ValueError("source or manifest hash mismatch")
        inspect(source, False)
        for name, expected in data["captures"].items():
            if digest((ROOT / data["artifactDirectory"] / name).read_bytes()) != expected:
                raise ValueError(f"capture hash mismatch: {name}")


if __name__ == "__main__":
    if sys.argv[1:] not in ([], ["--local"]):
        raise SystemExit("usage: verify-corpus-charts-powerpoint.py [--local]")
    verify(bool(sys.argv[1:]))
    print("verified charts package" + (" and local evidence" if sys.argv[1:] else ""))
