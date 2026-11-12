"""Compare PowerPoint and LibreOffice line records for equal authored text bodies."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

root = Path(__file__).resolve().parents[2]
output = Path(next((arg for arg in sys.argv[1:] if arg != "--record"), root / "docs/testing/tab-shared-source-apps-evidence.json"))
power_path = root / "docs/testing/tab-stop-wrap-matrix-powerpoint-evidence.json"
lo_path = root / "docs/testing/tab-position-libreoffice-evidence.json"
power, lo = json.loads(power_path.read_text()), json.loads(lo_path.read_text())
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
for verifier in ("verify-native-tab-wrap-matrix.py", "verify-tab-position-libreoffice.py"):
    subprocess.run(["python3", str(root / "scripts/research" / verifier)], cwd=root, check=True, capture_output=True)
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
source_a, source_b = root / power["directory"] / "matrix.pptx", root / lo["directory"] / "positions.pptx"
cases = []
with zipfile.ZipFile(source_a) as a, zipfile.ZipFile(source_b) as b:
    for name, power_slide, lo_slide in [("center", 6, 4), ("decimal", 8, 6)]:
        ra, rb = ET.fromstring(a.read(f"ppt/slides/slide{power_slide}.xml")), ET.fromstring(b.read(f"ppt/slides/slide{lo_slide}.xml"))
        def shape(root_node, identity):
            return next(item for item in root_node.iter(p + "sp") if item.find(p + "nvSpPr/" + p + "cNvPr").get("name") == identity)
        for index in range(3):
            sa, sb = shape(ra, f"{name}-150-{index}"), shape(rb, f"{name}-120-{index}")
            assert ET.tostring(sa.find(p + "txBody")) == ET.tostring(sb.find(p + "txBody"))
            assert ET.tostring(sa.find(p + "spPr")) == ET.tostring(sb.find(p + "spPr"))
            power_row, lo_row = power["observations"][power_slide - 1]["rows"][index], lo["pages"][lo_slide - 1]["rows"][index]
            lines = lambda rows: ["".join(char["char"] for char in line["chars"]) for line in rows]
            power_lines, lo_lines = lines(power_row), lines(lo_row)
            assert "".join(power_lines) == "".join(lo_lines)
            cases.append({"align": name, "sample": index, "text": "".join(power_lines), "powerpointLines": power_lines, "libreofficeLines": lo_lines, "sameLineGrouping": power_lines == lo_lines})
assert len(cases) == 6 and sum(not case["sameLineGrouping"] for case in cases) == 3
result = {"scope": "Six 150-pt-frame, 120-pt-stop samples with equal authored text bodies and geometry. LibreOffice input removes explicit zero indent; font candidates and PDF export differ, so no single-cause attribution or native-edit claim.", "powerpointEvidenceSha256": sha(power_path), "libreofficeEvidenceSha256": sha(lo_path), "verifierSha256": sha(Path(__file__)), "cases": cases}
if "--record" in sys.argv: output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
else: assert json.loads(output.read_text()) == result
print(json.dumps({"verified": True, "samples": len(cases), "differentLineGroupings": sum(not case["sameLineGrouping"] for case in cases)}))
