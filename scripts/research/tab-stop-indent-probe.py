"""Isolate paragraph attributes in a bounded LibreOffice headless PDF control."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

import fitz

root = Path(__file__).resolve().parents[2]
base = root / ".tmp/native-tab-stops"
base.mkdir(parents=True, exist_ok=True)
directory = Path(tempfile.mkdtemp(prefix="indent-", dir=base))
source = (root / sys.argv[1]).resolve()
if not source.is_relative_to(base.resolve()): raise ValueError("Source must be retained under .tmp/native-tab-stops")
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
shutil.copyfile(source, directory / "original.pptx")
names = ["original", "remove-indent", "remove-marL", "remove-algn", "remove-bullet", "positive-zero"]
for name in names[1:]:
    with zipfile.ZipFile(source) as z, zipfile.ZipFile(directory / f"{name}.pptx", "x", zipfile.ZIP_DEFLATED) as out:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename.startswith("ppt/slides/slide") and info.filename.endswith(".xml"):
                slide = ET.fromstring(data)
                for p in slide.iter(a + "pPr"):
                    if p.find(a + "tabLst") is None: continue
                    if name.startswith("remove-") and name != "remove-bullet": p.attrib.pop(name[7:], None)
                    if name == "remove-bullet":
                        for bullet in p.findall(a + "buNone"): p.remove(bullet)
                    if name == "positive-zero":
                        if p.get("indent") != "-0": raise ValueError("Expected negative lexical zero")
                        p.set("indent", "0")
                data = ET.tostring(slide, encoding="utf-8", xml_declaration=True)
            out.writestr(info, data)
args = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(directory), *[str(directory / f"{name}.pptx") for name in names]]
subprocess.run(args, check=True, capture_output=True, text=True)
labels = [["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["1.2", "123.45", "0.678"]]
evidence = {"scope": "One LibreOfficeDev headless conversion per variant; attribute isolation, not PowerPoint or native editing.",
            "application": subprocess.check_output(["soffice", "--version"], text=True).strip(),
            "runnerSha256": sha(Path(__file__)), "variants": []}
for name in names:
    presentation, pdf = directory / f"{name}.pptx", directory / f"{name}.pdf"
    pages = []
    with fitz.open(pdf) as doc:
        if len(doc) != len(labels): raise ValueError("Unexpected page count")
        for page, expected in zip(doc, labels):
            words = page.get_text("words"); row = []
            for text in expected:
                matches = [word for word in words if word[4] == text]
                if len(matches) != 1: raise ValueError("Missing or duplicate sample word")
                row.append({"text": text, "bbox": [round(float(v), 4) for v in matches[0][:4]]})
            pages.append(row)
    evidence["variants"].append({"name": name, "source": str(presentation.relative_to(root)), "sourceSha256": sha(presentation), "pdf": str(pdf.relative_to(root)), "pdfSha256": sha(pdf), "pages": pages})
print(json.dumps(evidence, indent=2, ensure_ascii=False))
