"""Change only the linked layout direction of the inheritance control."""
from pathlib import Path
from posixpath import join, normpath
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

folder = Path(sys.argv[1] if len(sys.argv) > 1 else ".tmp/vertical-writing")
source = folder / "vertical-placeholder-inherit.pptx"
destination = folder / "vertical-placeholder-layout-changed.pptx"
with ZipFile(source) as before:
    layouts = []
    for page in (1, 2):
        rels = ET.fromstring(before.read(f"ppt/slides/_rels/slide{page}.xml.rels"))
        target, = [rel.get("Target") for rel in rels if rel.get("Type", "").endswith("/slideLayout")]
        layouts.append(normpath(join("ppt/slides", target)).lstrip("/"))
    assert layouts[0] == layouts[1]
    layout = layouts[0]
    xml = before.read(layout)
    assert xml.count(b'vert="eaVert"') == 1
    with ZipFile(destination, "x") as after:
        for entry in before.infolist():
            data = xml.replace(b'vert="eaVert"', b'vert="vert270"') if entry.filename == layout else before.read(entry)
            after.writestr(entry, data)
