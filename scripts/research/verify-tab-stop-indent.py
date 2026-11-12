"""Verify that each retained variant changes only the stated paragraph property."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = (root / ".tmp/native-tab-stops").resolve()
evidence = json.loads((Path(sys.argv[1]) if len(sys.argv) > 1 else root / "docs/testing/tab-stop-indent-libreoffice-evidence.json").read_text())
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
if evidence["runnerSha256"] != sha(root / "scripts/research/tab-stop-indent-probe.py"): raise ValueError("Producer changed")
names = ["original", "remove-indent", "remove-marL", "remove-algn", "remove-bullet", "positive-zero"]
if [entry["name"] for entry in evidence["variants"]] != names: raise ValueError("Unexpected variants")
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
labels = [["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["1.2", "123.45", "0.678"]]
paths = {}
for variant in evidence["variants"]:
    paths[variant["name"]] = {}
    for name in ("source", "pdf"):
        file = (root / variant[name]).resolve()
        if not file.is_relative_to(base) or sha(file) != variant[name + "Sha256"]: raise ValueError("Missing or changed artifact")
        paths[variant["name"]][name] = file

def tree(node):
    return (node.tag, tuple(sorted(node.attrib.items())), node.text or "", node.tail or "", tuple(tree(child) for child in node))

positions = {}
with zipfile.ZipFile(paths["original"]["source"]) as original:
    original_names = original.namelist()
    if len(original_names) != len(set(original_names)): raise ValueError("Duplicate original parts")
    for variant in evidence["variants"]:
        name = variant["name"]
        with zipfile.ZipFile(paths[name]["source"]) as package:
            if package.namelist() != original_names: raise ValueError("Package part set changed")
            for part in original_names:
                if not (part.startswith("ppt/slides/slide") and part.endswith(".xml")):
                    if original.read(part) != package.read(part): raise ValueError("Unrelated part changed")
                    continue
                expected, actual = ET.fromstring(original.read(part)), ET.fromstring(package.read(part))
                for p in expected.iter(a + "pPr"):
                    if p.find(a + "tabLst") is None: continue
                    if name.startswith("remove-") and name != "remove-bullet": p.attrib.pop(name[7:], None)
                    if name == "remove-bullet":
                        for bullet in p.findall(a + "buNone"): p.remove(bullet)
                    if name == "positive-zero":
                        if p.get("indent") != "-0": raise ValueError("Original indent changed")
                        p.set("indent", "0")
                if tree(expected) != tree(actual): raise ValueError("Unexpected slide modification")
        with fitz.open(paths[name]["pdf"]) as pdf:
            if len(pdf) != 4: raise ValueError("Unexpected page count")
            pages = []
            for page, words_expected, recorded in zip(pdf, labels, variant["pages"]):
                words = page.get_text("words"); observed = []
                for text in words_expected:
                    matches = [word for word in words if word[4] == text]
                    if len(matches) != 1: raise ValueError("Missing sample word")
                    observed.append({"text": text, "bbox": [round(float(v), 4) for v in matches[0][:4]]})
                if observed != recorded: raise ValueError("PDF observation changed")
                pages.append([entry["bbox"][0] for entry in observed])
            positions[name] = pages
for name in ("remove-marL", "remove-algn", "remove-bullet", "positive-zero"):
    if positions[name] != positions["original"]: raise ValueError("Expected unchanged control")
if not all(abs(x - 132) < 1 for page in positions["original"] for x in page): raise ValueError("Original observation changed")
if not all(abs(x - 340) < 1 for x in positions["remove-indent"][0]): raise ValueError("Left-aligned control changed")
if not positions["remove-indent"][1][0] > positions["remove-indent"][1][1] > positions["remove-indent"][1][2]: raise ValueError("Center control changed")
print(json.dumps({"verified": True, "variants": len(names), "pagesPerVariant": 4}))
