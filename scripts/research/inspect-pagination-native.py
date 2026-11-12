"""Independently check the retained first-open native pagination observation."""
import hashlib
import json
import re
import sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

import fitz

root = Path(sys.argv[1])
source = root / "source.pptx"
patched = root / "style-probe.pptx"
pdfs = [root / "pptkit-pagination-probe-20261111.pdf", root / "style-probe.pdf"]
ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
counts = []
with ZipFile(source) as original, ZipFile(patched) as probe:
    assert set(original.namelist()) == set(probe.namelist())
    for name in original.namelist():
        before, after = original.read(name), probe.read(name)
        if re.fullmatch(r"ppt/slides/slide[1-5]\.xml", name):
            expected, count = re.subn(rb"<a:tcPr(?:\s[^>]*)?>.*?</a:tcPr>", b"<a:tcPr/>", before, flags=re.S)
            assert after == expected
            counts.append(count)
            table = ET.fromstring(after).find(".//a:tbl", ns)
            props = table.find("a:tblPr", ns)
            assert [props.get(key) for key in ("firstRow", "bandRow", "lastRow")] == ["1", "1", "1"]
        else:
            assert before == after
assert counts == [5, 2, 2, 2, 2]

observed = []
for pdf in pdfs:
    document = fitz.open(pdf)
    assert len(document) == 5
    colors = []
    for index, page in enumerate(document):
        assert page.rect.width == 400 and page.rect.height == 400
        text = page.get_text()
        expected_rows = range(5) if index == 0 else [0, index]
        assert text.split() == [part for row in expected_rows for part in ("ROW", str(row))]
        raster = page.get_pixmap(alpha=False)
        rows = []
        for y in ([80, 120, 160, 200, 240] if index == 0 else [80, 120]):
            samples = [list(raster.pixel(x, y)) for x in (60, 160, 260)]
            assert samples[0] == samples[1] == samples[2]
            rows.append(samples[1])
        colors.append(rows)
    observed.append(colors)
assert all(color == [255, 255, 255] for page in observed[0] for color in page)
assert observed[1][0] == [[237, 125, 49], [248, 215, 205], [252, 236, 232], [248, 215, 205], [237, 125, 49]]
assert all(page == [[237, 125, 49], [237, 125, 49]] for page in observed[1][1:])
def sha(file): return hashlib.sha256(file.read_bytes()).hexdigest()
print(json.dumps({"sourceSha256": sha(source), "patchedSha256": sha(patched), "pdfSha256": [sha(p) for p in pdfs], "patchedCells": counts, "sampleColors": observed}, indent=2))
