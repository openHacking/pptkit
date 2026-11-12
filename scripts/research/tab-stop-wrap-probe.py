"""Synthetic ODF controls for a single tab and a narrow text frame."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
base = root / ".tmp/native-tab-stops"
source = (root / sys.argv[1]).resolve()
if not source.is_relative_to(base.resolve()): raise ValueError("Source must be under .tmp/native-tab-stops")
directory = Path(tempfile.mkdtemp(prefix="wrap-", dir=base))
S = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
T = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
D = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
FO = "{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}"
V = "{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}"
cases = [("fit300", 300, "wrap", False), ("soft150", 150, "wrap", False), ("nowrap150", 150, "no-wrap", False), ("plain150", 150, "wrap", True)]
for name, width, wrap, plain in cases:
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(directory / f"{name}.odp", "x", zipfile.ZIP_DEFLATED) as output:
        for info in original.infolist():
            data = original.read(info.filename)
            if info.filename == "content.xml":
                document = ET.fromstring(data)
                shapes = [s for s in document.iter(D + "custom-shape") if s.find(".//" + T + "tab") is not None]
                if len(shapes) != 12: raise ValueError("Unexpected sample shapes")
                graphic_styles = {s.get(D + "style-name") for s in shapes}
                for shape in shapes:
                    shape.set(V + "width", f"{width}pt")
                    if plain:
                        for p in shape.iter(T + "p"):
                            for tab in p.iter(T + "tab"): tab.tag = T + "s"; tab.attrib.clear()
                            p.findall(T + "span")[-1].text = "WIDE WIDE"
                for style in document.iter(S + "style"):
                    name_attr = style.get(S + "name") or ""
                    if name_attr in graphic_styles: style.find(S + "graphic-properties").set(FO + "wrap-option", wrap)
                    if name_attr.startswith("Probe"):
                        style.find(S + "paragraph-properties/" + S + "tab-stops/" + S + "tab-stop").set(S + "position", "120pt")
                data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
            output.writestr(info, data)
subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(directory), *[str(directory / f"{name}.odp") for name, _, _, _ in cases]], check=True, capture_output=True, text=True)
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
result = {"scope": "Synthetic narrow ODF controls, one LibreOfficeDev headless PDF conversion per variant; not PowerPoint or native editing.",
          "application": subprocess.check_output(["soffice", "--version"], text=True).strip(), "runnerSha256": sha(Path(__file__)), "cases": []}
labels = [["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["I", "WIDE", "MIXED"], ["1.2", "123.45", "0.678"]]
for name, width, wrap, plain in cases:
    odp, pdf = directory / f"{name}.odp", directory / f"{name}.pdf"; pages = []
    with fitz.open(pdf) as doc:
        if len(doc) != 4: raise ValueError("Unexpected PDF pages")
        for page, expected in zip(doc, labels):
            allowed = {"P", "WIDE"} if plain else {"P", *expected}
            words = [word for word in page.get_text("words") if word[4] in allowed]
            if len(words) != (9 if plain else 6): raise ValueError("Missing or duplicated sample words")
            pages.append([{"text": word[4], "bbox": [round(float(v), 4) for v in word[:4]]} for word in words])
    result["cases"].append({"name": name, "width": width, "wrap": wrap, "plain": plain, "source": str(odp.relative_to(root)), "sourceSha256": sha(odp), "pdf": str(pdf.relative_to(root)), "pdfSha256": sha(pdf), "pages": pages})
print(json.dumps(result, ensure_ascii=False, indent=2))
