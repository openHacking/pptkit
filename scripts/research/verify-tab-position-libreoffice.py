"""Verify one LibreOffice print of the tab-position matrix with removed zero indents."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
output = root / "docs/testing/tab-position-libreoffice-evidence.json"
if "--record" in sys.argv:
    base = Path(sys.argv[sys.argv.index("--record") + 1]).resolve()
else:
    output = Path(sys.argv[1]) if len(sys.argv) > 1 else output
    base = (root / json.loads(output.read_text())["directory"]).resolve()
assert base.is_relative_to((root / ".tmp/native-tab-stops").resolve())
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
source, variant, pdf = base / "positions.pptx", base / "no-indent.pptx", base / "no-indent.pdf"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
controls = [("center", "ctr", ["I WORD", "WIDE WORD", "MIXED WORD"]), ("right", "r", ["I WORD", "WIDE WORD", "MIXED WORD"]), ("decimal", "dec", ["1.2 WORD", "123.45 WORD", "0.678 WORD"])]
stops = [90, 120, 180, 240]
assert json.loads((base / "manifest.json").read_text())["sha256"] == sha(source)
with zipfile.ZipFile(source) as original, zipfile.ZipFile(variant) as changed:
    assert set(original.namelist()) == set(changed.namelist())
    for filename in original.namelist():
        if not (filename.startswith("ppt/slides/slide") and filename.endswith(".xml") and "/_rels/" not in filename):
            assert original.read(filename) == changed.read(filename); continue
        before = ET.fromstring(original.read(filename)); after = ET.fromstring(changed.read(filename))
        for shape in before.iter(p + "sp"):
            if shape.find(".//" + a + "tabLst/" + a + "tab") is not None:
                for prop in shape.iter(a + "pPr"):
                    assert prop.attrib.pop("indent") == "-0"
        assert ET.tostring(before) == ET.tostring(after)
    for page_index, (position, (name, alignment, samples)) in enumerate(((position, control) for position in stops for control in controls), 1):
        slide = ET.fromstring(changed.read(f"ppt/slides/slide{page_index}.xml"))
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
        for sample_index, sample in enumerate(samples):
            start = 100 + 130 * sample_index; lines = {}
            for block in page.get_text("rawdict")["blocks"]:
                for line in block.get("lines", []):
                    if not start <= line["bbox"][1] < start + 130: continue
                    for span in line["spans"]:
                        for char in span["chars"]:
                            if char["c"].isspace(): continue
                            y = round(line["bbox"][1]); lines.setdefault(y, []).append({"char": char["c"], "bbox": [round(float(v), 4) for v in char["bbox"]]})
            assert "".join(char["char"] for chars in lines.values() for char in chars) == "P" + sample.replace(" ", "")
            rows.append([{"y": y, "chars": chars} for y, chars in lines.items()])
        pages.append({"align": name, "position": position, "rows": rows})
result = {"scope": "LibreOfficeDev headless conversion of a PPTX matrix with only sample paragraph indent attributes removed; no PowerPoint, editing, or reopen observation.", "application": subprocess.check_output(["soffice", "--version"], text=True).strip(), "directory": str(base.relative_to(root)), "sourceSha256": sha(source), "variantSha256": sha(variant), "pdfSha256": sha(pdf), "generatorSha256": sha(root / "scripts/prepare-native-tab-stop-position-matrix.mjs"), "producerSha256": sha(root / "scripts/research/probe-tab-position-libreoffice.py"), "verifierSha256": sha(Path(__file__)), "pages": pages}
if "--record" in sys.argv: output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
else: assert json.loads(output.read_text()) == result
print(json.dumps({"verified": True, "slides": len(pages), "samples": sum(len(page["rows"]) for page in pages)}))
