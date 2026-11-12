#!/usr/bin/env python3
"""Check generated shape controls against the retained native text-box source.

This establishes input equivalence only; it does not establish application output.
"""
import argparse
import hashlib
import json
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
CONTROLS = {"left": ["I", "WIDE", "MIXED"], "center": ["I", "WIDE", "MIXED"], "right": ["I", "WIDE", "MIXED"], "decimal": ["1.2", "123.45", "0.678"]}


def check(condition, message):
    if not condition:
        raise ValueError(message)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def shapes(archive, slide):
    root = ET.fromstring(archive.read(f"ppt/slides/slide{slide}.xml"))
    result = {}
    for item in root.findall(".//p:sp", NS):
        name = item.find("p:nvSpPr/p:cNvPr", NS).get("name")
        check(name not in result, f"duplicate shape name: {name}")
        result[name] = item
    return result


def verify(shape_file, shape_manifest, reference_file, reference_evidence):
    manifest = json.loads(Path(shape_manifest).read_text())
    reference = json.loads(Path(reference_evidence).read_text())
    check(sha(shape_file) == manifest["sha256"], "shape source hash mismatch")
    check(sha(reference_file) == reference["sourceSha256"], "reference source hash mismatch")
    check(manifest["position"] == 280 and manifest["controls"] == [{"align": k, "words": v} for k, v in CONTROLS.items()], "control manifest mismatch")
    compared = []
    with ZipFile(shape_file) as shape_zip, ZipFile(reference_file) as reference_zip:
        check(shape_zip.testzip() is None and reference_zip.testzip() is None, "corrupt package")
        for slide, (align, words) in enumerate(CONTROLS.items(), 1):
            shape_items, reference_items = shapes(shape_zip, slide), shapes(reference_zip, slide)
            for index, word in enumerate(words):
                name = f"{align}-{index}"
                shape, text = shape_items[name], reference_items[name]
                check(shape.find("p:nvSpPr/p:cNvSpPr", NS).get("txBox") is None, f"{name}: shape marked as text box")
                check(text.find("p:nvSpPr/p:cNvSpPr", NS).get("txBox") == "1", f"{name}: reference not marked as text box")
                shape_body, text_body = shape.find("p:txBody", NS), text.find("p:txBody", NS)
                check(ET.tostring(shape_body) == ET.tostring(text_body), f"{name}: text body differs")
                check([node.text for node in shape_body.findall(".//a:t", NS)] == ["P\t", word], f"{name}: text differs")
                shape_props, text_props = shape.find("p:spPr", NS), text.find("p:spPr", NS)
                check(ET.tostring(shape_props.find("a:xfrm", NS)) == ET.tostring(text_props.find("a:xfrm", NS)), f"{name}: geometry differs")
                check(shape_props.find("a:prstGeom", NS).get("prst") == "rect", f"{name}: not rectangular")
                fill = shape_props.find("a:solidFill/a:srgbClr", NS)
                check(fill is not None and fill.get("val") == "F2F5FA", f"{name}: fill differs")
                check(text_props.find("a:noFill", NS) is not None, f"{name}: reference is not unfilled")
                compared.append(name)
    return {"scope": "12 generated shape/text-box inputs: matching text bodies and transforms; shape fill differs. No native shape display, editing or reopening observation.", "shapeSha256": sha(shape_file), "referenceSha256": sha(reference_file), "items": compared}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--shape", default=".tmp/native-tab-stops/powerpoint-shape-j8Dn0A/shapes.pptx")
    parser.add_argument("--manifest", default=".tmp/native-tab-stops/powerpoint-shape-j8Dn0A/manifest.json")
    parser.add_argument("--reference", default=".tmp/native-tab-stops/run-snt5VZ/tab-stops.pptx")
    parser.add_argument("--evidence", default="docs/testing/tab-stop-powerpoint-evidence.json")
    args = parser.parse_args()
    print(json.dumps(verify(args.shape, args.manifest, args.reference, args.evidence)))
