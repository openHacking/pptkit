"""Verify one recorded PowerPoint first-open and print-to-PDF observation."""
import hashlib
import json
from pathlib import Path
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = root / ".tmp/native-tab-stops/run-snt5VZ"
manifest = Path(next((arg for arg in sys.argv[1:] if arg != "--record"), root / "docs/testing/tab-stop-powerpoint-evidence.json"))
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
source, pdf = base / "tab-stops.pptx", base / "tab-stops.pdf"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
alignments = [("left", "l", ["I", "WIDE", "MIXED"]), ("center", "ctr", ["I", "WIDE", "MIXED"]), ("right", "r", ["I", "WIDE", "MIXED"]), ("decimal", "dec", ["1.2", "123.45", "0.678"])]
with zipfile.ZipFile(source) as package:
    for index, (name, alignment, samples) in enumerate(alignments, 1):
        slide = ET.fromstring(package.read(f"ppt/slides/slide{index}.xml"))
        shapes = {shape.find(p + "nvSpPr/" + p + "cNvPr").get("name"): shape for shape in slide.iter(p + "sp")}
        for number, sample in enumerate(samples):
            shape = shapes[f"{name}-{number}"]
            para = shape.find(p + "txBody/" + a + "p")
            tab = para.find(a + "pPr/" + a + "tabLst/" + a + "tab")
            assert tab.attrib == {"pos": "3556000", "algn": alignment}
            assert [run.find(a + "t").text for run in para.findall(a + "r")] == ["P\t", sample]
            assert para.find(a + "pPr").get("indent") == "-0"
records = []
with fitz.open(pdf) as document:
    assert len(document) == 4
    for page_index, (page, (name, _, samples)) in enumerate(zip(document, alignments), 1):
        guides = [item for item in page.get_drawings() if item["fill"] and item["fill"][0] > 0.7 and item["fill"][1] < 0.3]
        assert len(guides) == 1
        guide = (guides[0]["rect"].x0 + guides[0]["rect"].x1) / 2
        words = [word for word in page.get_text("words") if word[4] in {"P", *samples}]
        assert [word[4] for word in words] == [value for sample in samples for value in ("P", sample)]
        dots = [char["bbox"][0] for block in page.get_text("rawdict")["blocks"] if "lines" in block for line in block["lines"] for span in line["spans"] for char in span["chars"] if char["c"] == "."]
        rows = []
        for number, sample in enumerate(samples):
            prefix, word = words[number * 2:number * 2 + 2]
            anchor = word[0] if name == "left" else (word[0] + word[2]) / 2 if name == "center" else word[2] if name == "right" else dots[number]
            assert abs(anchor - guide) < 0.25, (name, sample, anchor, guide)
            assert abs(prefix[1] - word[1]) < 0.25
            rows.append({"text": sample, "prefix": [round(x, 4) for x in prefix[:4]], "word": [round(x, 4) for x in word[:4]], "anchor": round(anchor, 4)})
        ax = base / f"native-slide-{page_index}.ax.txt"; text = ax.read_text()
        assert f"title-{name}" in text and all(f"{name}-{number}" in text for number in range(3))
        records.append({"slide": name, "guide": round(guide, 4), "rows": rows, "screenshotSha256": sha(base / f"native-slide-{page_index}.png"), "axSha256": sha(ax)})
rounds = []
for round_number in range(1, 4):
    saved = base / f"native-round-{round_number}.pptx"
    with zipfile.ZipFile(saved) as package:
        for index, (name, alignment, samples) in enumerate(alignments, 1):
            slide = ET.fromstring(package.read(f"ppt/slides/slide{index}.xml"))
            shapes = {shape.find(p + "nvSpPr/" + p + "cNvPr").get("name"): shape for shape in slide.iter(p + "sp")}
            for number, sample in enumerate(samples):
                para = shapes[f"{name}-{number}"].find(p + "txBody/" + a + "p")
                assert para.find(a + "pPr/" + a + "tabLst/" + a + "tab").attrib == {"pos": "3556000", "algn": alignment}
                assert "".join(node.text or "" for node in para.iter(a + "t")) == "P\t" + sample
    ax = base / f"native-round-{round_number}-reopen.ax.txt"
    assert f"native-round-{round_number}.pptx" in ax.read_text() and "left-1" in ax.read_text()
    rounds.append({"pptxSha256": sha(saved), "axSha256": sha(ax), "screenshotSha256": sha(base / f"native-round-{round_number}-reopen.png")})
edited = base / "native-edited.pptx"
with zipfile.ZipFile(edited) as package:
    for index, (name, alignment, samples) in enumerate(alignments, 1):
        slide = ET.fromstring(package.read(f"ppt/slides/slide{index}.xml"))
        shapes = {shape.find(p + "nvSpPr/" + p + "cNvPr").get("name"): shape for shape in slide.iter(p + "sp")}
        for number, sample in enumerate(samples):
            para = shapes[f"{name}-{number}"].find(p + "txBody/" + a + "p")
            assert para.find(a + "pPr/" + a + "tabLst/" + a + "tab").attrib == {"pos": "3556000", "algn": alignment}
            assert "".join(node.text or "" for node in para.iter(a + "t")) == "P\t" + sample + ("X" if name == "left" and number == 1 else "")
edit_ax = base / "native-edited-reopen.ax.txt"
assert "native-edited.pptx" in edit_ax.read_text() and "P\tWIDEX" in edit_ax.read_text()
assert "native-edited.pptx" in (base / "native-edited-before-close.ax.txt").read_text() and "P\tWIDEX" in (base / "native-edited-before-close.ax.txt").read_text()
edit = {"pptxSha256": sha(edited), "axSha256": sha(edit_ax), "screenshotSha256": sha(base / "native-edited-reopen.png"), "beforeCloseAxSha256": sha(base / "native-edited-before-close.ax.txt")}
version = plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())["CFBundleShortVersionString"]
result = {"scope": "PowerPoint first open and A4 scaled print-to-PDF of four generated slides; three Save As/close/reopen rounds check source XML on all slides and the first slide's window. A separate one-character native edit is saved and reopened once. No reopened PDF comparison. Screenshot and AX hashes are not independent proof of provenance.", "applicationVersion": version, "directory": str(base.relative_to(root)), "sourceSha256": sha(source), "pdfSha256": sha(pdf), "generatorSha256": sha(root / "scripts/prepare-native-tab-stops.mjs"), "verifierSha256": sha(Path(__file__)), "slides": records, "rounds": rounds, "edit": edit}
if "--record" in sys.argv: manifest.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
else: assert json.loads(manifest.read_text()) == result
print(json.dumps({"verified": True, "slides": len(records), "samples": sum(len(r["rows"]) for r in records)}))
