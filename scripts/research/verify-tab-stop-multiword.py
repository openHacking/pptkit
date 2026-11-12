"""Verify the retained ODF controls, their prior source and PDF word geometry."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
import fitz

root = Path(__file__).resolve().parents[2]
evidence = json.loads(Path(sys.argv[1] if len(sys.argv) > 1 else root / "docs/testing/tab-stop-multiword-libreoffice-evidence.json").read_text())
prior_path = root / "docs/testing/tab-stop-wrap-libreoffice-evidence.json"
prior = json.loads(prior_path.read_text())
sha = lambda data: hashlib.sha256(data).hexdigest()
assert sha(prior_path.read_bytes()) == evidence["priorSha256"]
assert sha((root / "scripts/research/tab-stop-multiword-probe.py").read_bytes()) == evidence["runnerSha256"]
subprocess.run(["python3", str(root / "scripts/research/verify-tab-stop-wrap.py")], cwd=root, check=True, capture_output=True)
assert [c["name"] for c in evidence["cases"]] == ["fit300", "soft150", "nowrap150"]
D = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
T = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
base = (root / ".tmp/native-tab-stops").resolve()
for case in evidence["cases"]:
    original = next(item for item in prior["cases"] if item["name"] == case["name"])
    odp, pdf = (root / case["source"]).resolve(), (root / case["pdf"]).resolve()
    assert odp.is_relative_to(base) and pdf.is_relative_to(base)
    assert sha(odp.read_bytes()) == case["sourceSha256"] and sha(pdf.read_bytes()) == case["pdfSha256"]
    with zipfile.ZipFile(root / original["source"]) as before, zipfile.ZipFile(odp) as after:
        assert set(before.namelist()) == set(after.namelist())
        for name in before.namelist():
            if name != "content.xml": assert before.read(name) == after.read(name); continue
            source, changed = ET.fromstring(before.read(name)), ET.fromstring(after.read(name))
            shapes = [shape for shape in changed.iter(D + "custom-shape") if shape.find(".//" + T + "tab") is not None]
            assert len(shapes) == 12
            for shape in shapes:
                span = next(shape.iter(T + "p")).findall(T + "span")[-1]
                assert span.text.endswith(" WORD"); span.text = span.text[:-5]
            assert ET.tostring(source) == ET.tostring(changed)
    with fitz.open(pdf) as doc:
        assert len(doc) == len(case["pages"]) == 4
        for page_index, page in enumerate(doc):
            if page_index == 2: assert case["pages"][page_index] == []; continue
            records = case["pages"][page_index]; assert len(records) == 3
            labels = ["I", "WIDE", "MIXED"] if page_index < 3 else ["1.2", "123.45", "0.678"]
            for index, (record, first) in enumerate(zip(records, labels)):
                start = 120 + 85 * index
                words = [w for w in page.get_text("words") if start <= w[1] < start + 85 and w[4] in {"P", first, "WORD"}]
                actual = [{"text": w[4], "bbox": [round(float(x), 4) for x in w[:4]]} for w in words]
                assert actual == record and [w["text"] for w in record] == ["P", first, "WORD"]
print(json.dumps({"verified": True, "cases": 3, "samples": 27}))
