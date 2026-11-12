"""Verify one PowerPoint narrow-frame first-open and print-to-PDF observation."""
import hashlib
import json
from pathlib import Path
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = root / ".tmp/native-tab-stops/powerpoint-wrap-vdAZ8A"
manifest = Path(next((arg for arg in sys.argv[1:] if arg != "--record"), root / "docs/testing/tab-stop-wrap-powerpoint-evidence.json"))
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
source, pdf = base / "tab-wrap.pptx", base / "tab-wrap.pdf"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
controls = [("left", "l", ["I WORD", "WIDE WORD", "MIXED WORD"]), ("center", "ctr", ["I WORD", "WIDE WORD", "MIXED WORD"]), ("right", "r", ["I", "WIDE", "MIXED"]), ("decimal", "dec", ["1.2 WORD", "123.45 WORD", "0.678 WORD"])]
expected = [[['PI', 'WORD'], ['P', 'W', 'IDEWORD'], ['P', 'M', 'IXED', 'WORD']], [['PIWORD'], ['PWIDE', 'WORD'], ['PMIXED', 'WORD']], [['PI'], ['PWIDE'], ['PMIXED']], [['P1.2', 'WORD'], ['P123.45', 'WORD'], ['P0.678', 'WORD']]]
with zipfile.ZipFile(source) as package:
    for index, (name, alignment, samples) in enumerate(controls, 1):
        slide = ET.fromstring(package.read(f"ppt/slides/slide{index}.xml"))
        shapes = {shape.find(p + "nvSpPr/" + p + "cNvPr").get("name"): shape for shape in slide.iter(p + "sp")}
        for number, sample in enumerate(samples):
            shape = shapes[f"{name}-{number}"]
            ext = shape.find(p + "spPr/" + a + "xfrm/" + a + "ext")
            assert ext.get("cx") == str(150 * 12700)
            body = shape.find(p + "txBody")
            assert body.find(a + "bodyPr").get("wrap") == "square"
            para = body.find(a + "p")
            assert para.find(a + "pPr/" + a + "tabLst/" + a + "tab").attrib == {"pos": str(120 * 12700), "algn": alignment}
            assert [run.find(a + "t").text for run in para.findall(a + "r")] == ["P\t", sample]
observations = []
with fitz.open(pdf) as document:
    assert len(document) == 4
    guides = []
    for index, page in enumerate(document):
        red = [drawing for drawing in page.get_drawings() if drawing["fill"] and drawing["fill"][0] > 0.7 and drawing["fill"][1] < 0.3]
        assert len(red) == 1
        guides.append((red[0]["rect"].x0 + red[0]["rect"].x1) / 2)
        rows = []
        for sample in range(3):
            start = 160 + 103 * sample; lines = {}
            for block in page.get_text("rawdict")["blocks"]:
                for line in block.get("lines", []):
                    if not start <= line["bbox"][1] < start + 103: continue
                    for span in line["spans"]:
                        for char in span["chars"]:
                            if char["c"].isspace(): continue
                            y = round(line["bbox"][1]); lines.setdefault(y, []).append({"char": char["c"], "bbox": [round(float(v), 4) for v in char["bbox"]]})
            assert ["".join(char["char"] for char in row) for row in lines.values()] == expected[index][sample]
            rows.append([{"y": y, "chars": chars} for y, chars in lines.items()])
        observations.append(rows)
    for index in (1, 2):
        rows = observations[0][index]
        assert abs(rows[1]["chars"][0]["bbox"][0] - guides[0]) < 0.25
        assert abs(rows[2]["chars"][0]["bbox"][0] - rows[0]["chars"][0]["bbox"][0]) < 0.25
    for index in range(3):
        dot = next(char for row in observations[3][index] for char in row["chars"] if char["char"] == ".")
        distance = abs(dot["bbox"][0] - guides[3])
        assert distance < 0.25 if index == 0 else distance > 4
assert "left-1" in (base / "native-slide-1.ax.txt").read_text()
version = plistlib.loads(Path("/Applications/Microsoft PowerPoint.app/Contents/Info.plist").read_bytes())["CFBundleShortVersionString"]
result = {"scope": "PowerPoint first open and scaled A4 print of four narrow tabbed slides; no save/reopen or native-edit claim. The screenshot and AX capture are not independent provenance proof.", "applicationVersion": version, "directory": str(base.relative_to(root)), "sourceSha256": sha(source), "pdfSha256": sha(pdf), "screenshotSha256": sha(base / "native-slide-1.png"), "axSha256": sha(base / "native-slide-1.ax.txt"), "generatorSha256": sha(root / "scripts/prepare-native-tab-wrap.mjs"), "verifierSha256": sha(Path(__file__)), "guides": [round(guide, 4) for guide in guides], "observations": observations}
if "--record" in sys.argv: manifest.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
else: assert json.loads(manifest.read_text()) == result
print(json.dumps({"verified": True, "slides": 4, "samples": 12}))
