"""Recheck an eight-page PowerPoint first-open, scaled-print phrase matrix."""
import hashlib
import json
from pathlib import Path
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = root / ".tmp/native-tab-stops/powerpoint-phrases-xJFZ3n"
manifest = Path(next((arg for arg in sys.argv[1:] if arg != "--record"), root / "docs/testing/tab-phrases-powerpoint-evidence.json"))
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
source, pdf = base / "phrases.pptx", base / "phrases.pdf"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
controls = [("left", "l", ["I WIDE WORD", "WIDE WORD MORE", "MIXED WORD LONG"]), ("center", "ctr", ["I WIDE WORD", "WIDE WORD MORE", "MIXED WORD LONG"]), ("right", "r", ["I WIDE WORD", "WIDE WORD MORE", "MIXED WORD LONG"]), ("decimal", "dec", ["1.2 WIDE WORD", "123.45 WORD MORE", "0.678 WORD LONG"])]
widths = [150, 220]
source_data = json.loads((base / "manifest.json").read_text())
assert source_data["sha256"] == sha(source) and source_data["widths"] == widths and source_data["position"] == 120
with zipfile.ZipFile(source) as package:
    for page_index, (width, (name, alignment, samples)) in enumerate(((width, control) for width in widths for control in controls), 1):
        slide = ET.fromstring(package.read(f"ppt/slides/slide{page_index}.xml"))
        shapes = {shape.find(p + "nvSpPr/" + p + "cNvPr").get("name"): shape for shape in slide.iter(p + "sp")}
        for number, sample in enumerate(samples):
            shape = shapes[f"{name}-{width}-{number}"]; ext = shape.find(p + "spPr/" + a + "xfrm/" + a + "ext")
            assert ext.get("cx") == str(width * 12700) and ext.get("cy") == str(180 * 12700)
            body = shape.find(p + "txBody"); assert body.find(a + "bodyPr").get("wrap") == "square"
            para = body.find(a + "p")
            assert para.find(a + "pPr/" + a + "tabLst/" + a + "tab").attrib == {"pos": str(120 * 12700), "algn": alignment}
            assert [run.find(a + "t").text for run in para.findall(a + "r")] == ["P\t", sample]
observations = []
with fitz.open(pdf) as document:
    assert len(document) == 8
    for page_index, page in enumerate(document):
        name, _, samples = controls[page_index % 4]; width = widths[page_index // 4]
        red = [drawing for drawing in page.get_drawings() if drawing["fill"] and drawing["fill"][0] > 0.7 and drawing["fill"][1] < 0.3]
        assert len(red) == 1; guide = (red[0]["rect"].x0 + red[0]["rect"].x1) / 2
        rows = []
        for sample_index, sample in enumerate(samples):
            start = 160 + 103 * sample_index; lines = {}
            for block in page.get_text("rawdict")["blocks"]:
                for line in block.get("lines", []):
                    if not start <= line["bbox"][1] < start + (130 if sample_index == 2 else 103): continue
                    for span in line["spans"]:
                        for char in span["chars"]:
                            if char["c"].isspace(): continue
                            y = round(line["bbox"][1]); lines.setdefault(y, []).append({"char": char["c"], "bbox": [round(float(v), 4) for v in char["bbox"]]})
            assert "".join(char["char"] for chars in lines.values() for char in chars) == "P" + sample.replace(" ", "")
            rows.append([{"y": y, "chars": chars} for y, chars in lines.items()])
        observations.append({"width": width, "align": name, "guide": round(guide, 4), "rows": rows})
def strings(page, row): return ["".join(char["char"] for char in line["chars"]) for line in observations[page]["rows"][row]]
assert strings(1, 0) == ["PIWIDE", "WORD"]
assert strings(5, 1) == ["PWIDEWORD", "MORE"]
assert strings(3, 0) == ["P1.2WIDE", "WORD"]
assert strings(0, 2) == ["P", "M", "IXED", "WORD", "LONG"]
assert "left-150-1" in (base / "first-open.ax.txt").read_text()
version = plistlib.loads(Path("/Applications/Microsoft PowerPoint.app/Contents/Info.plist").read_bytes())["CFBundleShortVersionString"]
result = {"scope": "PowerPoint first open and scaled A4 print of an eight-page three-item phrase matrix; no native edit or save/reopen. Screenshot and AX are not independent provenance proof.", "applicationVersion": version, "directory": str(base.relative_to(root)), "sourceSha256": sha(source), "sourceManifestSha256": sha(base / "manifest.json"), "pdfSha256": sha(pdf), "screenshotSha256": sha(base / "first-open.png"), "axSha256": sha(base / "first-open.ax.txt"), "generatorSha256": sha(root / "scripts/prepare-native-tab-phrases.mjs"), "verifierSha256": sha(Path(__file__)), "observations": observations}
if "--record" in sys.argv: manifest.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
else: assert json.loads(manifest.read_text()) == result
print(json.dumps({"verified": True, "slides": len(observations), "samples": sum(len(page["rows"]) for page in observations)}))
