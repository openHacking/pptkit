"""Check first-open native column counts and matching frame/cell ink."""
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from PIL import Image, ImageChops

folder = Path(sys.argv[1])
ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main", "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
with ZipFile(folder / "east-asian-wrap-control.pptx") as archive:
    slide = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
    shapes = slide.findall(".//p:sp", ns)
    cells = slide.findall(".//a:tc", ns)
    labels = ["".join(n.text for n in s.findall(".//a:t", ns)) for s in (shapes[0], shapes[2], shapes[4], shapes[5])]
    assert labels == ["FRAME 180", "FRAME 55", "CELL 180", "CELL 55"]
    for shape, height in zip((shapes[1], shapes[3]), (180, 55), strict=True):
        assert shape.find("p:spPr/a:xfrm/a:ext", ns).get("cy") == str(height * 12700)
        assert shape.find("p:txBody/a:bodyPr", ns).get("vert") == "eaVert"
    assert [row.get("h") for row in slide.findall(".//a:tr", ns)] == [str(n * 12700) for n in (180, 55)]
    assert [c.find("a:tcPr", ns).get("vert") for c in cells] == ["eaVert", "eaVert"]

with Image.open(folder / "east-asian-wrap-native.png") as image:
    assert image.size == (1920, 1080)
    im = image.convert("RGB")
crops = []
for x, expected in zip((120, 520, 920, 1320), (1, 3, 1, 3), strict=True):
    crop = im.crop((x, 220, x + 300, 580))
    mask = crop.convert("L").point(lambda v: 255 if v < 96 else 0)
    occupied = [bool(mask.crop((i, 0, i + 1, mask.height)).getbbox()) for i in range(mask.width)]
    columns = sum(ink and (i == 0 or not occupied[i - 1]) for i, ink in enumerate(occupied))
    assert columns == expected, (x, columns, expected)
    crops.append(crop)
for i in (0, 1):
    assert ImageChops.difference(crops[i], crops[i + 2]).getbbox() is None
print("180 pt: one column; 55 pt: three columns; frame and cell crops match")
