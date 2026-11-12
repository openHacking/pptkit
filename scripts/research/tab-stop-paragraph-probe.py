"""Create ODF paragraph controls and observe one LibreOffice PDF conversion each."""
import copy
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
source = (root / sys.argv[1]).resolve()
if not source.is_relative_to(base.resolve()): raise ValueError("Source must be retained under .tmp/native-tab-stops")
directory = Path(tempfile.mkdtemp(prefix="paragraph-", dir=base))
shutil.copyfile(source, directory / "input.pptx")
subprocess.run(["soffice", "--headless", "--convert-to", "odp", "--outdir", str(directory), str(directory / "input.pptx")], check=True, capture_output=True, text=True)
S = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
T = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
D = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
O = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
FO = "{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}"
cases = [("zero", 0, 0), ("left36", 36, 0), ("hang18", 36, -18), ("start0", 18, -18)]
for name, margin, indent in cases:
    with zipfile.ZipFile(directory / "input.odp") as source_zip, zipfile.ZipFile(directory / f"{name}.odp", "x", zipfile.ZIP_DEFLATED) as output:
        for info in source_zip.infolist():
            data = source_zip.read(info.filename)
            if info.filename == "content.xml":
                document = ET.fromstring(data); styles = document.find(O + "automatic-styles")
                original = next(node for node in styles if node.get(S + "name") == "P2")
                pages = list(document.iter(D + "page"))
                if len(pages) != 4: raise ValueError("Unexpected slide count")
                for index, (page, align) in enumerate(zip(pages, ["left", "center", "right", "char"])):
                    style = copy.deepcopy(original); style.set(S + "name", f"Probe{index}")
                    props = style.find(S + "paragraph-properties"); props.set(FO + "margin-left", f"{margin}pt"); props.set(FO + "text-indent", f"{indent}pt")
                    stops = props.find(S + "tab-stops"); stops.clear()
                    stop = ET.SubElement(stops, S + "tab-stop", {S + "position": "9.877777cm", S + "type": align})
                    if align == "char": stop.set(S + "char", ".")
                    styles.append(style)
                    selected = [p for p in page.iter(T + "p") if p.find(".//" + T + "tab") is not None]
                    if len(selected) != 3: raise ValueError("Unexpected sample paragraphs")
                    for p in selected: p.set(T + "style-name", f"Probe{index}")
                data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
            output.writestr(info, data)
subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(directory), *[str(directory / f"{name}.odp") for name, _, _ in cases]], check=True, capture_output=True, text=True)
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
labels = [["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["1.2", "123.45", "0.678"]]
evidence = {"scope": "Synthetic ODF controls, one LibreOfficeDev headless PDF conversion per variant; not PowerPoint or native editing.", "application": subprocess.check_output(["soffice", "--version"], text=True).strip(), "runnerSha256": sha(Path(__file__)), "cases": []}
for name, margin, indent in cases:
    odp, pdf = directory / f"{name}.odp", directory / f"{name}.pdf"; pages = []
    with fitz.open(pdf) as doc:
        if len(doc) != 4: raise ValueError("Unexpected PDF pages")
        for page, expected in zip(doc, labels):
            words = page.get_text("words"); rows = []
            for label in expected:
                match = [word for word in words if word[4] == label]
                if len(match) != 1: raise ValueError("Missing sample")
                rows.append({"text": label, "bbox": [round(float(v), 4) for v in match[0][:4]]})
            prefixes = [word for word in words if word[4] == "P"]
            if len(prefixes) != 3: raise ValueError("Missing prefix")
            pages.append({"samples": rows, "prefixX": [round(float(word[0]), 4) for word in prefixes]})
    evidence["cases"].append({"name": name, "margin": margin, "indent": indent, "source": str(odp.relative_to(root)), "sourceSha256": sha(odp), "pdf": str(pdf.relative_to(root)), "pdfSha256": sha(pdf), "pages": pages})
print(json.dumps(evidence, ensure_ascii=False, indent=2))
