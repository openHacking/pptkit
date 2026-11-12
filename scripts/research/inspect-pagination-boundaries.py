"""Check one first-open source-boundary pagination research control."""
import hashlib
import json
import re
import sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

import fitz

root = Path(sys.argv[1])
source, patched, pdf = (root / name for name in ("source.pptx", "style-probe-boundaries.pptx", "style-probe-boundaries.pdf"))
ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
seen = {}
with ZipFile(source) as original, ZipFile(patched) as control:
    assert set(original.namelist()) == set(control.namelist())
    for name in original.namelist():
        before, after = original.read(name), control.read(name)
        match = re.fullmatch(r"ppt/slides/slide([1-5])\.xml", name)
        if not match:
            assert before == after
            continue
        index = int(match.group(1))
        expected, count = re.subn(rb"<a:tcPr(?:\s[^>]*)?>.*?</a:tcPr>", b"<a:tcPr/>", before, flags=re.S)
        assert after == expected and count == (5 if index == 1 else 2)
        props = ET.fromstring(after).find(".//a:tbl/a:tblPr", ns)
        assert props.get("firstRow") == "1" and props.get("bandRow") == "1"
        assert props.get("lastRow") == ("1" if index in (1, 5) else None)
        seen[index] = count
assert seen == {1: 5, 2: 2, 3: 2, 4: 2, 5: 2}
document = fitz.open(pdf)
assert len(document) == 5
colors = []
for index, page in enumerate(document):
    assert page.rect.width == 400 and page.rect.height == 400
    expected_rows = range(5) if index == 0 else [0, index]
    assert page.get_text().split() == [part for row in expected_rows for part in ("ROW", str(row))]
    pix = page.get_pixmap(alpha=False)
    rows = []
    for y in ([80, 120, 160, 200, 240] if index == 0 else [80, 120]):
        samples = [list(pix.pixel(x, y)) for x in (60, 160, 260)]
        assert samples[0] == samples[1] == samples[2]
        rows.append(samples[0])
    colors.append(rows)
orange, band_a, band_b = [237, 125, 49], [248, 215, 205], [252, 236, 232]
assert colors == [[orange, band_a, band_b, band_a, orange], [orange, band_a], [orange, band_a], [orange, band_a], [orange, orange]]
sha = lambda file: hashlib.sha256(file.read_bytes()).hexdigest()
print(json.dumps({"sourceSha256": sha(source), "patchedSha256": sha(patched), "pdfSha256": sha(pdf), "colors": colors}, indent=2))
