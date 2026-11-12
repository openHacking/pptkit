"""Verify retained input packages and PDF observations independently of the producer."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = (root / ".tmp/native-tab-stops").resolve()
evidence = json.loads((Path(sys.argv[1]) if len(sys.argv) > 1 else root / "docs/testing/tab-stops-libreoffice-evidence.json").read_text())
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
for key, name in [("generatorSha256", "scripts/prepare-native-tab-stops.mjs"), ("runnerSha256", "scripts/research/probe-tab-stops.py")]:
    if sha(root / name) != evidence[key]: raise ValueError("Producer changed")
if [v["name"] for v in evidence["variants"]] != ["tab-stops", "same-run", "no-stops"]: raise ValueError("Unexpected variant set")
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
labels = [["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["1.2", "123.45", "0.678"]]
positions = []
for variant in evidence["variants"]:
    paths = []
    for name in ("source", "pdf"):
        file = (root / variant[name]).resolve()
        if not file.is_relative_to(base) or sha(file) != variant[name + "Sha256"]: raise ValueError("Missing or changed artifact")
        paths.append(file)
    with zipfile.ZipFile(paths[0]) as package:
        if len(package.namelist()) != len(set(package.namelist())): raise ValueError("Duplicate ZIP entries")
        for index, align in enumerate(("l", "ctr", "r", "dec"), 1):
            slide = ET.fromstring(package.read(f"ppt/slides/slide{index}.xml"))
            actual = [node.attrib for node in slide.iter(a + "tab")]
            expected = [] if variant["name"] == "no-stops" else [{"pos": "3556000", "algn": align}] * 3
            if actual != expected: raise ValueError("Unexpected tab-stop declarations")
            paragraphs = [p for p in slide.iter(a + "p") if any("P\t" in (t.text or "") for t in p.iter(a + "t"))]
            if len(paragraphs) != 3: raise ValueError("Missing sample paragraphs")
            for p in paragraphs:
                runs = p.findall(a + "r")
                if len(runs) != (1 if variant["name"] == "same-run" else 2): raise ValueError("Run boundary changed")
    with fitz.open(paths[1]) as pdf:
        if len(pdf) != 4: raise ValueError("Unexpected PDF page count")
        pages = []
        for page, expected, recorded in zip(pdf, labels, variant["pages"]):
            words = page.get_text("words")
            observed = []
            for word in expected:
                candidates = [w for w in words if w[4] == word]
                if len(candidates) != 1: raise ValueError("Missing or duplicated PDF word")
                observed.append({"text": word, "bbox": [round(float(v), 4) for v in candidates[0][:4]]})
            if observed != recorded: raise ValueError("PDF text geometry changed")
            pages.append([entry["bbox"][0] for entry in observed])
        positions.append(pages)
if positions[0] != positions[1] or positions[0] != positions[2]: raise ValueError("Control positions differ")
if not all(abs(x - 132) < 1 for page in positions[0] for x in page): raise ValueError("Observation no longer matches documented scope")
print(json.dumps({"verified": True, "variants": 3, "pagesPerVariant": 4, "observedLeftX": positions[0]}))
