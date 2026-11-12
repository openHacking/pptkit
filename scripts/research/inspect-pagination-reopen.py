"""Check one PowerPoint save/reopen of the actual source-boundary output."""
import hashlib
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import fitz

root = Path(sys.argv[1])
source, saved, before, reopened = (root / name for name in ("source.pptx", "native-saved.pptx", "native-saved.pdf", "native-reopened.pdf"))
ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main", "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
expected_rows = [[f"ROW {i}" for i in range(5)]] + [["ROW 0", f"ROW {i}"] for i in range(1, 5)]
for file in (source, saved):
    with ZipFile(file) as package:
        for index, expected in enumerate(expected_rows, 1):
            page = ET.fromstring(package.read(f"ppt/slides/slide{index}.xml"))
            table = page.find(".//a:tbl", ns)
            assert [node.text for node in table.findall(".//a:t", ns)] == expected
            props = table.find("a:tblPr", ns)
            assert props.get("firstRow") == "1" and props.get("bandRow") == "1"
            assert props.get("lastRow") == ("1" if index in (1, 5) else None)
            assert [row.get("h") for row in table.findall("a:tr", ns)] == ["508000"] * len(expected)
            transform = page.find(".//p:graphicFrame/p:xfrm", ns)
            assert transform.find("a:off", ns).attrib == {"x": "508000", "y": "762000"}
            assert transform.find("a:ext", ns).attrib == {"cx": "3048000", "cy": "2540000" if index == 1 else "1016000"}
docs = [fitz.open(before), fitz.open(reopened)]
assert all(len(doc) == 5 for doc in docs)
for index, (first, second) in enumerate(zip(*docs)):
    assert first.rect == second.rect == fitz.Rect(0, 0, 400, 400)
    expected = [part for row in expected_rows[index] for part in row.split()]
    assert first.get_text().split() == second.get_text().split() == expected
    assert first.get_pixmap(alpha=False).samples == second.get_pixmap(alpha=False).samples
sha = lambda file: hashlib.sha256(file.read_bytes()).hexdigest()
print(json.dumps({"sourceSha256": sha(source), "savedSha256": sha(saved), "postSavePdfSha256": sha(before), "reopenedPdfSha256": sha(reopened), "pages": 5, "equalRasterPages": 5}, indent=2))
