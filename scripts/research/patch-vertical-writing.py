"""Prepare a schema-mode matrix while leaving non-slide parts untouched."""
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

folder = Path(sys.argv[1] if len(sys.argv) > 1 else ".tmp/vertical-writing")
ns = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
for prefix, uri in ns.items():
    ET.register_namespace(prefix, uri)
modes = ["horz", "vert", "vert270", "eaVert", "mongolianVert", "wordArtVert", "wordArtVertRtl"]
targets = {f"ppt/slides/slide{i}.xml": mode for i, mode in enumerate(modes, 1)}
with ZipFile(folder / "horizontal.pptx") as original, ZipFile(folder / "vertical-matrix.pptx", "w") as output:
    assert set(targets) <= set(original.namelist())
    for entry in original.infolist():
        data = original.read(entry)
        if entry.filename in targets:
            mode = targets[entry.filename]
            slide = ET.fromstring(data)
            boxes = slide.findall(".//p:sp/p:txBody", ns)
            cells = slide.findall(".//a:tc", ns)
            assert len(boxes) == len(cells) == 2
            assert boxes[0].findtext(".//a:t", namespaces=ns) == "MODE: " + mode
            boxes[1].find("a:bodyPr", ns).set("vert", mode)
            cells[0].find("a:tcPr", ns).set("vert", mode)
            cells[1].find("a:txBody/a:bodyPr", ns).set("vert", mode)
            data = ET.tostring(slide, encoding="utf-8", xml_declaration=True)
        output.writestr(entry, data)
