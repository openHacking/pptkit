"""One headless LibreOffice PDF observation of generated list/tab controls."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = (root / ".tmp/native-tab-list").resolve()
directory = Path(sys.argv[1]).resolve()
if not directory.is_relative_to(base): raise ValueError("Expected a retained native-tab-list directory")
source, pdf = directory / "list-tabs.pptx", directory / "list-tabs.pdf"
manifest = json.loads((directory / "manifest.json").read_text())
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
if sha(source) != manifest["sourceSha256"] or len(manifest["cases"]) != 36: raise ValueError("Source manifest mismatch")
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
with zipfile.ZipFile(source) as z:
    if len(z.namelist()) != len(set(z.namelist())): raise ValueError("Duplicate ZIP entries")
    for case in manifest["cases"]:
        slide = ET.fromstring(z.read(f'ppt/slides/slide{case["index"] + 1}.xml'))
        p = next(p for p in slide.iter(a + "p") if p.find(a + "pPr/" + a + "tabLst") is not None)
        pr, tab = p.find(a + "pPr"), p.find(a + "pPr/" + a + "tabLst/" + a + "tab")
        if int(pr.get("marL")) != case["indent"] * 12700 or int(pr.get("indent")) != -case["hanging"] * 12700: raise ValueError("Unexpected indent")
        if tab.attrib != {"pos": "3556000", "algn": {"left": "l", "center": "ctr", "right": "r", "decimal": "dec"}[case["align"]]}: raise ValueError("Unexpected tab stop")
        if [node.text for node in p.iter(a + "t")] != ["P\t", case["text"]]: raise ValueError("Unexpected source text")
        if case["bullet"] == "symbol":
            if p.find(a + "pPr/" + a + "buChar").get("char") != "•": raise ValueError("Unexpected bullet")
        elif int(p.find(a + "pPr/" + a + "buAutoNum").get("startAt")) != int(case["bullet"].removeprefix("number")): raise ValueError("Unexpected numbering")
subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(directory), str(source)], check=True, capture_output=True, text=True)
observations = []
with fitz.open(pdf) as doc:
    if len(doc) != len(manifest["cases"]): raise ValueError("Unexpected PDF pages")
    for page, case in zip(doc, manifest["cases"]):
        raw = page.get_text("rawdict"); chars = [(c, span["font"]) for block in raw["blocks"] if "lines" in block for line in block["lines"] for span in line["spans"] for c in span["chars"]]
        prefix = [(char, font) for char, font in chars if char["c"] == "P"]
        samples = [word for word in page.get_text("words") if word[4] == case["text"]]
        if len(prefix) != 1 or len(samples) != 1: raise ValueError("Sample text missing or duplicated")
        sample = samples[0]; marker = [{"text": char["c"], "font": font, "bbox": [round(float(v), 4) for v in char["bbox"]]} for char, font in chars if char["c"] not in "P \t" and char["bbox"][2] < prefix[0][0]["bbox"][0] + 0.1]
        observations.append({**case, "prefix": {"text": "P", "font": prefix[0][1], "bbox": [round(float(v), 4) for v in prefix[0][0]["bbox"]]}, "sample": {"text": case["text"], "bbox": [round(float(v), 4) for v in sample[:4]]}, "marker": marker})
evidence = {"scope": "36 generated PPTX list/tab controls, one LibreOfficeDev headless PDF conversion; no PowerPoint or editing claim.",
            "application": subprocess.check_output(["soffice", "--version"], text=True).strip(), "generatorSha256": sha(root / "scripts/prepare-tab-list-controls.mjs"), "runnerSha256": sha(Path(__file__)),
            "source": str(source.relative_to(root)), "sourceSha256": sha(source), "pdf": str(pdf.relative_to(root)), "pdfSha256": sha(pdf), "observations": observations}
print(json.dumps(evidence, ensure_ascii=False, indent=2))
