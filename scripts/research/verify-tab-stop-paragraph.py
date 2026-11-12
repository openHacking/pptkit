"""Verify ODF paragraph controls and independently re-read their PDF geometry."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = (root / ".tmp/native-tab-stops").resolve()
evidence = json.loads((Path(sys.argv[1]) if len(sys.argv) > 1 else root / "docs/testing/tab-stop-paragraph-libreoffice-evidence.json").read_text())
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
if evidence["runnerSha256"] != sha(root / "scripts/research/tab-stop-paragraph-probe.py"): raise ValueError("Producer changed")
cases = [("zero", 0, 0), ("left36", 36, 0), ("hang18", 36, -18), ("start0", 18, -18)]
if [(x["name"], x["margin"], x["indent"]) for x in evidence["cases"]] != cases: raise ValueError("Unexpected cases")
S = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
T = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
D = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
FO = "{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}"
labels = [["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["1.2", "123.45", "0.678"]]
def tree(node): return (node.tag, tuple(sorted(node.attrib.items())), node.text or "", node.tail or "", tuple(tree(child) for child in node))
reference_tree = None; reference_parts = None; all_pages = []
for case in evidence["cases"]:
    paths = []
    for key in ["source", "pdf"]:
        file = (root / case[key]).resolve()
        if not file.is_relative_to(base) or sha(file) != case[key + "Sha256"]: raise ValueError("Missing or changed artifact")
        paths.append(file)
    with zipfile.ZipFile(paths[0]) as package:
        names = package.namelist()
        if len(names) != len(set(names)): raise ValueError("Duplicate ZIP entries")
        parts = {name: sha_bytes for name in names if name != "content.xml" for sha_bytes in [hashlib.sha256(package.read(name)).hexdigest()]}
        if reference_parts is None: reference_parts = parts
        elif parts != reference_parts: raise ValueError("Unrelated ODF parts changed")
        document = ET.fromstring(package.read("content.xml"))
        for index, align in enumerate(["left", "center", "right", "char"]):
            style = next(n for n in document.iter(S + "style") if n.get(S + "name") == f"Probe{index}")
            props = style.find(S + "paragraph-properties")
            if props.get(FO + "margin-left") != f'{case["margin"]}pt' or props.get(FO + "text-indent") != f'{case["indent"]}pt': raise ValueError("Paragraph control changed")
            stop = props.find(S + "tab-stops/" + S + "tab-stop")
            if stop.get(S + "position") != "9.877777cm" or stop.get(S + "type") != align or (align == "char" and stop.get(S + "char") != "."): raise ValueError("Tab control changed")
            props.set(FO + "margin-left", "0pt"); props.set(FO + "text-indent", "0pt")
        pages = list(document.iter(D + "page"))
        if len(pages) != 4: raise ValueError("Unexpected ODF pages")
        for index, page in enumerate(pages):
            selected = [p for p in page.iter(T + "p") if p.find(".//" + T + "tab") is not None]
            if len(selected) != 3 or any(p.get(T + "style-name") != f"Probe{index}" for p in selected): raise ValueError("Control paragraph selection changed")
        normalized = tree(document)
        if reference_tree is None: reference_tree = normalized
        elif normalized != reference_tree: raise ValueError("Unexpected ODF content difference")
    measured = []
    with fitz.open(paths[1]) as pdf:
        if len(pdf) != 4: raise ValueError("Unexpected PDF pages")
        decimal_chars = [c for block in pdf[3].get_text("rawdict")["blocks"] if "lines" in block for line in block["lines"] for span in line["spans"] for c in span["chars"] if c["c"] == "." and c["bbox"][1] > 100]
        if len(decimal_chars) != 3 or any(abs(c["bbox"][0] - (340 + case["margin"])) > 0.1 for c in decimal_chars): raise ValueError("Decimal tab position changed")
        for page, expected, recorded in zip(pdf, labels, case["pages"]):
            words = page.get_text("words"); samples = []
            spans = [span for block in page.get_text("dict")["blocks"] if "lines" in block for line in block["lines"] for span in line["spans"]]
            for text in expected:
                matches = [word for word in words if word[4] == text]
                if len(matches) != 1: raise ValueError("Missing sample")
                selected = [span for span in spans if span["text"] == text]
                expected_font = "LiberationSans-Bold" if text in ("MIXED", "0.678") else "LiberationSans"
                if len(selected) != 1 or selected[0]["font"] != expected_font: raise ValueError("Unexpected PDF font")
                samples.append({"text": text, "bbox": [round(float(v), 4) for v in matches[0][:4]]})
            prefixes = [round(float(word[0]), 4) for word in words if word[4] == "P"]
            if {"samples": samples, "prefixX": prefixes} != recorded: raise ValueError("PDF geometry changed")
            measured.append(recorded)
    all_pages.append(measured)
for (name, margin, indent), pages in zip(cases, all_pages):
    for base_page, page in zip(all_pages[0], pages):
        for actual, baseline in zip(page["prefixX"], base_page["prefixX"]):
            if abs((actual - baseline) - (margin + indent)) > 0.1: raise ValueError("Prefix offset changed")
        for actual, baseline in zip(page["samples"], base_page["samples"]):
            if abs((actual["bbox"][0] - baseline["bbox"][0]) - margin) > 0.1: raise ValueError("Tab offset changed")
if any(abs(value - 60) > 0.1 for value in all_pages[0][0]["prefixX"]): raise ValueError("Baseline text origin changed")
if any(abs(item["bbox"][0] - 340) > 0.1 for item in all_pages[0][0]["samples"]): raise ValueError("Baseline left tab position changed")
if any(abs((item["bbox"][0] + item["bbox"][2]) / 2 - 340) > 0.1 for item in all_pages[0][1]["samples"]): raise ValueError("Baseline center tab position changed")
if any(abs(item["bbox"][2] - 340) > 0.1 for item in all_pages[0][2]["samples"]): raise ValueError("Baseline right tab position changed")
print(json.dumps({"verified": True, "cases": len(cases), "pagesPerCase": 4}))
