"""Check logical Mongolian text and font use in the PowerPoint PDF."""
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import fitz
from PIL import Image

folder = Path(sys.argv[1])
text = "ᠮᠣᠩᠭᠣᠯ"
ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
with ZipFile(folder / "mongolian-native.pptx") as package:
    root = ET.fromstring(package.read("ppt/slides/slide1.xml"))
    _, body = root.findall(".//p:sp/p:txBody", ns)
    assert body.find("a:bodyPr", ns).get("vert") == "mongolianVert"
    assert body.find(".//a:rPr/a:cs", ns).get("typeface") == "Noto Sans Mongolian"
    assert "".join(n.text or "" for n in body.findall(".//a:t", ns)) == text
with fitz.open(folder / "mongolian-native.pdf") as document:
    assert len(document) == 1
    span, = [s for s in document[0].get_texttrace() if any(0x1800 <= c[0] <= 0x18AF for c in s["chars"])]
    assert "".join(chr(c[0]) for c in span["chars"]) == text
    assert span["font"] in ("NotoSansMongolian-Regular", "NotoSansMongolian-Regula")
    assert span["dir"] == (0.0, 1.0)
    assert all(c[1] > 0 for c in span["chars"])
with Image.open(folder / "mongolian-native.png") as image:
    assert image.size == (1920, 1080)
    bounds = image.crop((110, 185, 520, 550)).convert("L").point(lambda v: 255 if v < 96 else 0).getbbox()
    assert bounds and bounds[3] - bounds[1] > bounds[2] - bounds[0]
print("native package, downward PDF font trace, and visible vertical ink verified")
