"""Independently verify generated list/tab controls and retained PDF geometry."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = (root / ".tmp/native-tab-list").resolve()
e = json.loads((Path(sys.argv[1]) if len(sys.argv) > 1 else root / "docs/testing/tab-list-libreoffice-evidence.json").read_text())
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
for key, name in [("generatorSha256", "scripts/prepare-tab-list-controls.mjs"), ("runnerSha256", "scripts/research/probe-tab-lists.py")]:
    if sha(root / name) != e[key]: raise ValueError("Producer changed")
files = []
for key in ["source", "pdf"]:
    file = (root / e[key]).resolve()
    if not file.is_relative_to(base) or sha(file) != e[key + "Sha256"]: raise ValueError("Missing or changed artifact")
    files.append(file)
expected = [(bullet, indent, hanging, align) for bullet in ["symbol", "number3", "number333"] for indent, hanging in [(36, 18), (27, 27), (54, 18)] for align in ["left", "center", "right", "decimal"]]
if [(x["bullet"], x["indent"], x["hanging"], x["align"]) for x in e["observations"]] != expected: raise ValueError("Unexpected control matrix")
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
with zipfile.ZipFile(files[0]) as package, fitz.open(files[1]) as pdf:
    if len(package.namelist()) != len(set(package.namelist())) or len(pdf) != len(expected): raise ValueError("Invalid package/page count")
    for index, (case, page) in enumerate(zip(e["observations"], pdf)):
        if case["index"] != index or case["text"] != ("12.34" if case["align"] == "decimal" else "WIDE"): raise ValueError("Unexpected sample")
        slide = ET.fromstring(package.read(f"ppt/slides/slide{index + 1}.xml"))
        paragraphs = [p for p in slide.iter(a + "p") if p.find(a + "pPr/" + a + "tabLst") is not None]
        if len(paragraphs) != 1: raise ValueError("Unexpected paragraph")
        p = paragraphs[0]; pr = p.find(a + "pPr"); tab = pr.find(a + "tabLst/" + a + "tab")
        if int(pr.get("marL")) != case["indent"] * 12700 or int(pr.get("indent")) != -case["hanging"] * 12700: raise ValueError("Indent changed")
        if tab.attrib != {"pos": "3556000", "algn": {"left": "l", "center": "ctr", "right": "r", "decimal": "dec"}[case["align"]]}: raise ValueError("Stop changed")
        if [node.text for node in p.iter(a + "t")] != ["P\t", case["text"]]: raise ValueError("Source text changed")
        if case["bullet"] == "symbol":
            marker_expected = "•"; node = pr.find(a + "buChar")
            if node is None or node.get("char") != marker_expected: raise ValueError("Bullet changed")
        else:
            number = int(case["bullet"].removeprefix("number")); marker_expected = f"{number}."; node = pr.find(a + "buAutoNum")
            if node is None or node.get("type") != "arabicPeriod" or int(node.get("startAt")) != number: raise ValueError("Numbering changed")
        raw = page.get_text("rawdict"); chars = [(c, span["font"]) for block in raw["blocks"] if "lines" in block for line in block["lines"] for span in line["spans"] for c in span["chars"]]
        prefix = [(char, font) for char, font in chars if char["c"] == "P"]
        words = [word for word in page.get_text("words") if word[4] == case["text"]]
        if len(prefix) != 1 or len(words) != 1: raise ValueError("Missing text")
        spans = [span for block in page.get_text("dict")["blocks"] if "lines" in block for line in block["lines"] for span in line["spans"]]
        sample_spans = [span for span in spans if span["text"] == case["text"]]
        if prefix[0][1] != "LiberationSans" or len(sample_spans) != 1 or sample_spans[0]["font"] != "LiberationSans": raise ValueError("Unexpected PDF text font")
        prefix_record = {"text": "P", "font": prefix[0][1], "bbox": [round(float(v), 4) for v in prefix[0][0]["bbox"]]}
        word_record = {"text": case["text"], "bbox": [round(float(v), 4) for v in words[0][:4]]}
        marker = [{"text": char["c"], "font": font, "bbox": [round(float(v), 4) for v in char["bbox"]]} for char, font in chars if char["c"] not in "P \t" and char["bbox"][2] < prefix[0][0]["bbox"][0] + 0.1]
        if prefix_record != case["prefix"] or word_record != case["sample"] or marker != case["marker"]: raise ValueError("PDF observation changed")
        if "".join(x["text"] for x in marker) != marker_expected: raise ValueError("Marker text changed")
        start = 60 + case["indent"] - case["hanging"]
        if abs(marker[0]["bbox"][0] - start) > 0.1: raise ValueError("Marker start changed")
        if abs(prefix_record["bbox"][0] - max(60 + case["indent"], marker[-1]["bbox"][2])) > 0.2: raise ValueError("Text start changed")
        x0, _, x1, _ = word_record["bbox"]
        if case["align"] == "left" and abs(x0 - 340) > 0.1: raise ValueError("Left stop changed")
        if case["align"] == "center" and abs((x0 + x1) / 2 - 340) > 0.1: raise ValueError("Center stop changed")
        if case["align"] == "right" and abs(x1 - 340) > 0.1: raise ValueError("Right stop changed")
        if case["align"] == "decimal":
            dots = [c for c, _ in chars if c["c"] == "." and c["bbox"][0] > 250]
            if len(dots) != 1 or abs(dots[0]["bbox"][0] - 340) > 0.1: raise ValueError("Decimal stop changed")
print(json.dumps({"verified": True, "cases": len(expected)}))
