"""Recheck PowerPoint's 12-page first-open tab-position matrix."""
import hashlib
import json
from pathlib import Path
import plistlib
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = root / ".tmp/native-tab-stops/powerpoint-position-J3HUEA"
output = Path(next((arg for arg in sys.argv[1:] if arg != "--record"), root / "docs/testing/tab-position-powerpoint-evidence.json"))
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
source, pdf = base / "positions.pptx", base / "positions.pdf"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
controls = [("center", "ctr", ["I WORD", "WIDE WORD", "MIXED WORD"]), ("right", "r", ["I WORD", "WIDE WORD", "MIXED WORD"]), ("decimal", "dec", ["1.2 WORD", "123.45 WORD", "0.678 WORD"])]
stops = [90, 120, 180, 240]
assert json.loads((base / "manifest.json").read_text())["sha256"] == sha(source)
with zipfile.ZipFile(source) as package:
    for page_index, (position, (name, alignment, samples)) in enumerate(((position, control) for position in stops for control in controls), 1):
        slide = ET.fromstring(package.read(f"ppt/slides/slide{page_index}.xml"))
        shapes = {shape.find(p + "nvSpPr/" + p + "cNvPr").get("name"): shape for shape in slide.iter(p + "sp")}
        for index, sample in enumerate(samples):
            shape = shapes[f"{name}-{position}-{index}"]
            assert shape.find(p + "spPr/" + a + "xfrm/" + a + "ext").get("cx") == str(150 * 12700)
            body = shape.find(p + "txBody"); assert body.find(a + "bodyPr").get("wrap") == "square"
            para = body.find(a + "p"); assert para.find(a + "pPr/" + a + "tabLst/" + a + "tab").attrib == {"pos": str(position * 12700), "algn": alignment}
            assert [run.find(a + "t").text for run in para.findall(a + "r")] == ["P\t", sample]
pages = []
with fitz.open(pdf) as document:
    assert len(document) == 12
    for page_index, page in enumerate(document):
        name, _, samples = controls[page_index % 3]; position = stops[page_index // 3]; rows = []
        red = [drawing for drawing in page.get_drawings() if drawing["fill"] and drawing["fill"][0] > 0.7 and drawing["fill"][1] < 0.3]
        assert len(red) == 1; guide = (red[0]["rect"].x0 + red[0]["rect"].x1) / 2
        for sample_index, sample in enumerate(samples):
            start = 160 + 103 * sample_index; lines = {}
            for block in page.get_text("rawdict")["blocks"]:
                for line in block.get("lines", []):
                    if not start <= line["bbox"][1] < start + 103: continue
                    for span in line["spans"]:
                        for char in span["chars"]:
                            if char["c"].isspace(): continue
                            y = round(line["bbox"][1]); lines.setdefault(y, []).append({"char": char["c"], "bbox": [round(float(v), 4) for v in char["bbox"]]})
            assert "".join(char["char"] for chars in lines.values() for char in chars) == "P" + sample.replace(" ", "")
            rows.append([{"y": y, "chars": chars} for y, chars in lines.items()])
        pages.append({"align": name, "position": position, "guide": round(guide, 4), "rows": rows})
def first_word_x(page, row):
    first = pages[page]["rows"][row][0]["chars"]
    return first[1]["bbox"][0] - first[0]["bbox"][0]
assert abs(first_word_x(1, 1) - first_word_x(1, 0)) < 0.1
assert first_word_x(7, 1) > first_word_x(4, 1) + 20
assert "center-90-1" in (base / "first-open.ax.txt").read_text()
version = plistlib.loads(Path("/Applications/Microsoft PowerPoint.app/Contents/Info.plist").read_bytes())["CFBundleShortVersionString"]
result = {"scope": "PowerPoint first-open and scaled A4 print of a 12-page tab-position matrix; no native edit or save/reopen. Screenshot and AX are not independent provenance proof.", "applicationVersion": version, "directory": str(base.relative_to(root)), "sourceSha256": sha(source), "pdfSha256": sha(pdf), "screenshotSha256": sha(base / "first-open.png"), "axSha256": sha(base / "first-open.ax.txt"), "generatorSha256": sha(root / "scripts/prepare-native-tab-stop-position-matrix.mjs"), "verifierSha256": sha(Path(__file__)), "pages": pages}
if "--record" in sys.argv: output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
else: assert json.loads(output.read_text()) == result
print(json.dumps({"verified": True, "slides": len(pages), "samples": sum(len(page["rows"]) for page in pages)}))
