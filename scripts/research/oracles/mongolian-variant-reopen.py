"""Verify Mongolian control characters and visible ink across a native save."""
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from PIL import Image, ImageChops

folder = Path(sys.argv[1])
expected = "ᠠ\u180B ᠠ\u180C ᠠ\u180D ᠠ\u180Eᠠ"
ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
for name in ("pre-save", "mongolian-native"):
    with ZipFile(folder / f"{name}.pptx") as package:
        root = ET.fromstring(package.read("ppt/slides/slide1.xml"))
        _, body = root.findall(".//p:sp/p:txBody", ns)
        assert "".join(n.text or "" for n in body.findall(".//a:t", ns)) == expected
        assert body.find("a:bodyPr", ns).get("vert") == "mongolianVert"
with Image.open(folder / "first.png") as first, Image.open(folder / "reopen.png") as reopen:
    assert first.size == reopen.size == (1920, 1080)
    region = (100, 180, 550, 850)
    a, b = (image.convert("RGB").crop(region) for image in (first, reopen))
    assert a.convert("L").getextrema()[0] < 96
    assert ImageChops.difference(a, b).getbbox() is None
print("variant controls and mode retained; nonblank native body unchanged")
