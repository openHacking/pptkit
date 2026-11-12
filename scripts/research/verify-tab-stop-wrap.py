"""Independently inspect the narrow ODF controls and retained PDF words."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = (root / ".tmp/native-tab-stops").resolve()
e = json.loads((Path(sys.argv[1]) if len(sys.argv) > 1 else root / "docs/testing/tab-stop-wrap-libreoffice-evidence.json").read_text())
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
if e["runnerSha256"] != sha(root / "scripts/research/tab-stop-wrap-probe.py"): raise ValueError("Producer changed")
cases = [("fit300", 300, "wrap", False), ("soft150", 150, "wrap", False), ("nowrap150", 150, "no-wrap", False), ("plain150", 150, "wrap", True)]
if [(x["name"], x["width"], x["wrap"], x["plain"]) for x in e["cases"]] != cases: raise ValueError("Unexpected matrix")
S = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
T = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
D = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
FO = "{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}"
V = "{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}"
labels = [["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["1.2", "123.45", "0.678"]]
for case in e["cases"]:
    files = []
    for key in ["source", "pdf"]:
        file = (root / case[key]).resolve()
        if not file.is_relative_to(base) or sha(file) != case[key + "Sha256"]: raise ValueError("Missing or changed artifact")
        files.append(file)
    with zipfile.ZipFile(files[0]) as package:
        if len(package.namelist()) != len(set(package.namelist())): raise ValueError("Duplicate ODF parts")
        document = ET.fromstring(package.read("content.xml"))
        shapes = [shape for shape in document.iter(D + "custom-shape") if (shape.get(D + "name") or "").startswith(("left-", "center-", "right-", "decimal-"))]
        if len(shapes) != 12: raise ValueError("Missing sample shapes")
        styles = {style.get(S + "name"): style for style in document.iter(S + "style")}
        for shape in shapes:
            if shape.get(V + "width") != f'{case["width"]}pt': raise ValueError("Frame width changed")
            graphic = styles[shape.get(D + "style-name")].find(S + "graphic-properties")
            if graphic.get(FO + "wrap-option") != case["wrap"]: raise ValueError("Wrap mode changed")
            paragraphs = list(shape.iter(T + "p"))
            if len(paragraphs) != 1: raise ValueError("Unexpected paragraph count")
            paragraph = paragraphs[0]; name = shape.get(D + "name"); alignment, ordinal = name.rsplit("-", 1)
            page = ["left", "center", "right", "decimal"].index(alignment); expected = labels[page][int(ordinal)]
            if case["plain"]:
                if list(paragraph.iter(T + "tab")) or paragraph.findall(T + "span")[-1].text != "WIDE WIDE": raise ValueError("Plain control changed")
            elif len(list(paragraph.iter(T + "tab"))) != 1 or paragraph.findall(T + "span")[-1].text != expected: raise ValueError("Tabbed text changed")
            style = styles[paragraph.get(T + "style-name")]; stop = style.find(S + "paragraph-properties/" + S + "tab-stops/" + S + "tab-stop")
            if stop.get(S + "position") != "120pt" or stop.get(S + "type") != ["left", "center", "right", "char"][page]: raise ValueError("Stop changed")
    with fitz.open(files[1]) as pdf:
        if len(pdf) != 4: raise ValueError("Unexpected PDF pages")
        for page, expected, recorded in zip(pdf, labels, case["pages"]):
            allowed = {"P", "WIDE"} if case["plain"] else {"P", *expected}
            words = [{"text": w[4], "bbox": [round(float(v), 4) for v in w[:4]]} for w in page.get_text("words") if w[4] in allowed]
            if words != recorded or len(words) != (9 if case["plain"] else 6): raise ValueError("PDF observation changed")
for index, page in enumerate(e["cases"][1]["pages"]):
    for row in range(3):
        prefix, sample = page[row * 2:row * 2 + 2]
        should_wrap = row > 0 and index != 2
        if should_wrap:
            if abs(sample["bbox"][0] - 60) > 0.1 or sample["bbox"][1] - prefix["bbox"][1] < 25: raise ValueError("Expected soft break missing")
        elif abs(sample["bbox"][1] - prefix["bbox"][1]) > 0.1: raise ValueError("Unexpected soft break")
for case in [e["cases"][0], e["cases"][2]]:
    for page in case["pages"]:
        for row in range(3):
            if abs(page[row * 2]["bbox"][1] - page[row * 2 + 1]["bbox"][1]) > 0.1: raise ValueError("Control unexpectedly wrapped")
for page in e["cases"][3]["pages"]:
    if page[2]["bbox"][1] - page[1]["bbox"][1] < 25: raise ValueError("Plain wrap control did not wrap")
print(json.dumps({"verified": True, "cases": len(cases), "pagesPerCase": 4}))
