#!/usr/bin/env python3
"""Verify the retained native table observation; local captures are optional."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / "docs/testing/corpus-table-powerpoint-evidence.json"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def inventory(path):
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("ppt/slides/slide1.xml"))
    tables = root.findall(".//a:tbl", NS)
    if len(tables) != 1:
        raise ValueError("expected one table")
    rows = tables[0].findall("a:tr", NS)
    cells = [[(cell.attrib, "".join(t.text or "" for t in cell.findall(".//a:t", NS)))
              for cell in row.findall("a:tc", NS)] for row in rows]
    if [len(row) for row in cells] != [3, 3, 3]:
        raise ValueError("unexpected table grid")
    return cells


def check_table(cells, edited):
    expected = [["Merged heading", "", ""], [edited, "A", "B"], ["", "C", "D"]]
    if [[cell[1] for cell in row] for row in cells] != expected:
        raise ValueError("unexpected cell text")
    for row, col, key, value in [(0, 0, "gridSpan", "3"), (0, 1, "hMerge", "1"),
                                 (0, 2, "hMerge", "1"), (1, 0, "rowSpan", "2"),
                                 (2, 0, "vMerge", "1")]:
        if cells[row][col][0].get(key) != value:
            raise ValueError(f"missing {key} at {row},{col}")


def verify(include_local=False, fresh=False):
    evidence = ROOT / "docs/testing/corpus-table-powerpoint-fresh-evidence.json" if fresh else EVIDENCE
    record = json.loads(evidence.read_text())
    output = ROOT / record["output"]["path"]
    if digest(output) != record["output"]["sha256"]:
        raise ValueError("output hash mismatch")
    check_table(inventory(output), "Edited rows")
    if include_local:
        source = ROOT / record["input"]["path"]
        manifest = json.loads((ROOT / record["sourceManifest"]).read_text())
        entry = next(item for item in manifest["fixtures"] if item["id"] == "table")
        if digest(source) != record["input"]["sha256"] or entry["sha256"] != record["input"]["sha256"]:
            raise ValueError("source or manifest hash mismatch")
        check_table(inventory(source), "Two rows")
        for name, expected in record["sha256"].items():
            if digest(ROOT / record["artifactDirectory"] / name) != expected:
                raise ValueError(f"capture hash mismatch: {name}")


if __name__ == "__main__":
    if any(arg not in ("--local", "--fresh") for arg in sys.argv[1:]) or len(set(sys.argv[1:])) != len(sys.argv[1:]):
        raise SystemExit("usage: verify-corpus-table-powerpoint.py [--fresh] [--local]")
    verify(include_local="--local" in sys.argv, fresh="--fresh" in sys.argv)
    print("verified native table package" + (" and local evidence" if "--local" in sys.argv else ""))
