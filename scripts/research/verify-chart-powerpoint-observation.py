#!/usr/bin/env python3
"""Recheck retained files and named XML facts, not manual PowerPoint observations."""
import hashlib
import json
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
NAMESPACES = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}


def verify():
    observation = json.loads((ROOT / "docs/testing/chart-value-powerpoint-observation.json").read_text())
    earlier = json.loads((ROOT / "docs/testing/chart-value-save-evidence.json").read_text())
    assert observation["source"]["sha256"] == earlier["source"]["sha256"]
    assert observation["candidate"]["sha256"] == earlier["rounds"][2]["savedSha256"]
    for kind in ("source", "candidate", "screenshot"):
        record = observation[kind]
        path = (ROOT / record["path"]).resolve(strict=True)
        assert path.is_relative_to(ROOT)
        assert hashlib.sha256(path.read_bytes()).hexdigest() == record["sha256"], kind
        if kind == "screenshot":
            continue
        with zipfile.ZipFile(path) as archive:
            chart = ET.fromstring(archive.read("ppt/charts/chart1.xml"))
        axis = chart.find(".//c:valAx/c:scaling", NAMESPACES)
        point = chart.find(".//c:barChart/c:ser/c:val/c:numRef/c:numCache/c:pt[@idx='1']/c:v", NAMESPACES)
        actual = {"axisMinimum": axis.find("c:min", NAMESPACES).get("val"),
                  "axisMaximum": axis.find("c:max", NAMESPACES).get("val"),
                  "selectedCacheValue": point.text}
        assert actual == record["chart"], kind
    assert observation["source"]["chart"] == {"axisMinimum": "0", "axisMaximum": "9", "selectedCacheValue": "5"}
    assert observation["candidate"]["chart"] == {"axisMinimum": "0", "axisMaximum": "9", "selectedCacheValue": "37.25"}
    print("Retained source, candidate and screenshot hashes and chart XML verified; UI observations are manual.")


if __name__ == "__main__":
    verify()
