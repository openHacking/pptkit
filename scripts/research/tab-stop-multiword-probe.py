"""Produce independent LibreOffice controls for tab-aligned multiword wrapping."""
import hashlib
import json
from pathlib import Path
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
prior_path = root / "docs/testing/tab-stop-wrap-libreoffice-evidence.json"
prior = json.loads(prior_path.read_text())
subprocess.run(["python3", str(root / "scripts/research/verify-tab-stop-wrap.py")], cwd=root, check=True, capture_output=True)
base = root / ".tmp/native-tab-stops"
directory = Path(tempfile.mkdtemp(prefix="multiword-", dir=base))
S = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
T = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
D = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
names = ("fit300", "soft150", "nowrap150")
for name in names:
    case = next(item for item in prior["cases"] if item["name"] == name)
    with zipfile.ZipFile(root / case["source"]) as original, zipfile.ZipFile(directory / f"{name}.odp", "x", zipfile.ZIP_DEFLATED) as output:
        for info in original.infolist():
            data = original.read(info.filename)
            if info.filename == "content.xml":
                document = ET.fromstring(data)
                shapes = [shape for shape in document.iter(D + "custom-shape") if shape.find(".//" + T + "tab") is not None]
                if len(shapes) != 12: raise ValueError("Unexpected sample shapes")
                for shape in shapes:
                    span = next(shape.iter(T + "p")).findall(T + "span")[-1]
                    span.text += " WORD"
                data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
            output.writestr(info, data)
subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(directory), *[str(directory / f"{name}.odp") for name in names]], check=True, capture_output=True, text=True)
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
result = {"scope": "Synthetic ODF multiword controls, one LibreOfficeDev headless PDF conversion per variant; not PowerPoint or native editing.", "application": subprocess.check_output(["soffice", "--version"], text=True).strip(), "priorSha256": sha(prior_path), "runnerSha256": sha(Path(__file__)), "cases": []}
for name in names:
    odp, pdf = directory / f"{name}.odp", directory / f"{name}.pdf"; pages = []
    with fitz.open(pdf) as doc:
        if len(doc) != 4: raise ValueError("Unexpected page count")
        for page_index, page in enumerate(doc):
            # Right-aligned controls overlap the prefix in this fixture and
            # PDF word extraction merges the runs; keep them outside this profile.
            if page_index == 2:
                pages.append([])
                continue
            records = []
            for sample_index, first in enumerate((["I", "WIDE", "MIXED"] if page_index < 3 else ["1.2", "123.45", "0.678"])):
                start = 120 + 85 * sample_index
                words = [w for w in page.get_text("words") if start <= w[1] < start + 85 and w[4] in {"P", first, "WORD"}]
                if [w[4] for w in words] != ["P", first, "WORD"]: raise ValueError(f"Unexpected words {name} {page_index} {sample_index}: {words}")
                records.append([{"text": w[4], "bbox": [round(float(x), 4) for x in w[:4]]} for w in words])
            pages.append(records)
    result["cases"].append({"name": name, "source": str(odp.relative_to(root)), "sourceSha256": sha(odp), "pdf": str(pdf.relative_to(root)), "pdfSha256": sha(pdf), "pages": pages})
print(json.dumps(result, ensure_ascii=False, indent=2))
