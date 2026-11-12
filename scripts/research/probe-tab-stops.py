"""One isolated LibreOffice PDF observation; not a PowerPoint oracle."""
import hashlib
import json
from pathlib import Path
import platform
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

import fitz

directory = Path(sys.argv[1]).resolve()
root_dir = Path(__file__).resolve().parents[2]
if not directory.is_relative_to(root_dir / ".tmp/native-tab-stops"):
    raise ValueError("Evidence directory must be inside .tmp/native-tab-stops")
source = directory / "tab-stops.pptx"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()


def make_variant(name, variant):
    output = directory / f"{name}.pptx"
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(output, "x", zipfile.ZIP_DEFLATED) as result:
        for item in original.infolist():
            data = original.read(item.filename)
            if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
                root = ET.fromstring(data)
                for paragraph in root.iter(a + "p"):
                    properties = paragraph.find(a + "pPr")
                    if variant == "no-stops" and properties is not None:
                        for tablist in properties.findall(a + "tabLst"):
                            properties.remove(tablist)
                    runs = paragraph.findall(a + "r")
                    if variant == "same-run" and len(runs) == 2 and runs[0].find(a + "t").text == "P\t":
                        runs[0].find(a + "t").text += runs[1].find(a + "t").text
                        paragraph.remove(runs[1])
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            result.writestr(item, data)
    return output


variants = [source, make_variant("same-run", "same-run"), make_variant("no-stops", "no-stops")]
expected = [["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["1.2", "123.45", "0.678"]]
result = {"scope": "Four slides, one headless LibreOffice PDF conversion per variant; not native editing or PowerPoint evidence.",
          "application": subprocess.check_output(["soffice", "--version"], text=True).strip(), "platform": platform.platform(),
          "generatorSha256": sha(root_dir / "scripts/prepare-native-tab-stops.mjs"), "runnerSha256": sha(Path(__file__)), "variants": []}
for item in variants:
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(directory), str(item)], check=True, capture_output=True, text=True)
    pdf = item.with_suffix(".pdf")
    with fitz.open(pdf) as document:
        assert len(document) == 4
        pages = []
        for page, labels in zip(document, expected):
            words = page.get_text("words")
            rows = []
            for label in labels:
                matching = [word for word in words if word[4] == label]
                assert len(matching) == 1, (item, label, matching)
                word = matching[0]
                rows.append({"text": label, "bbox": [round(float(v), 4) for v in word[:4]]})
            pages.append(rows)
    result["variants"].append({"name": item.stem, "source": str(item.relative_to(root_dir)), "sourceSha256": sha(item), "pdf": str(pdf.relative_to(root_dir)), "pdfSha256": sha(pdf), "pages": pages})
print(json.dumps(result, ensure_ascii=False, indent=2))
