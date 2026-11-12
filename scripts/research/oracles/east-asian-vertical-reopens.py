"""Verify the source owners and native captures of the East Asian sample."""
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from PIL import Image, ImageChops

folder = Path(sys.argv[1])
ns = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
regions = [(100, 55, 530, 145), (100, 155, 770, 850), (880, 155, 1270, 850), (1270, 155, 1690, 850)]

def text(node):
    return "".join(t.text or "" for t in node.findall(".//a:t", ns))

for suffix in ("pre-save", "saved-1", "saved-2", "saved-3"):
    with ZipFile(folder / f"east-asian-{suffix}.pptx") as archive:
        root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        label, frame = root.findall(".//p:sp/p:txBody", ns)
        vertical, horizontal = root.findall(".//a:tc", ns)
        assert text(label) == "MODE: eaVert"
        assert frame.find("a:bodyPr", ns).get("vert") == "eaVert"
        assert vertical.find("a:tcPr", ns).get("vert") == "eaVert"
        assert horizontal.find("a:tcPr", ns).get("vert") is None
        for cell in (vertical, horizontal):
            assert cell.find("a:txBody/a:bodyPr", ns).get("vert") is None
        assert [text(x) for x in (frame, vertical, horizontal)] == ["日本語（ABC 123）"] * 3
    print(suffix, "owners and logical text verified")
with Image.open(folder / "mode-4.png") as image:
    reference = image.convert("RGB")
assert reference.size == (1920, 1080)
for region in regions:
    assert reference.crop(region).convert("L").getextrema()[0] < 96
for suffix in ("first", "reopen-1", "reopen-2", "reopen-3"):
    with Image.open(folder / f"east-asian-{suffix}.png") as image:
        assert image.size == reference.size
        for region in regions:
            assert ImageChops.difference(reference.crop(region), image.convert("RGB").crop(region)).getbbox() is None, (suffix, region)
    print(suffix, "matched all four capture regions")

if "--edited" in sys.argv[2:]:
    with ZipFile(folder / "east-asian-edit.pptx") as archive:
        root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        label, frame = root.findall(".//p:sp/p:txBody", ns)
        vertical, horizontal = root.findall(".//a:tc", ns)
        assert [text(x) for x in (label, frame, vertical, horizontal)] == [
            "MODE: eaVert", "日本語（ABC 456）", "日本語（ABC 789）", "日本語（ABC 123）",
        ]
        assert frame.find("a:bodyPr", ns).get("vert") == "eaVert"
        assert vertical.find("a:tcPr", ns).get("vert") == "eaVert"
        assert horizontal.find("a:tcPr", ns).get("vert") is None
        for cell in (vertical, horizontal):
            assert cell.find("a:txBody/a:bodyPr", ns).get("vert") is None
    with Image.open(folder / "east-asian-edit-reopen.png") as image:
        assert image.size == reference.size
        actual = image.convert("RGB")
        for i, region in enumerate(regions):
            crop = actual.crop(region)
            assert crop.convert("L").getextrema()[0] < 96, i
            changed = ImageChops.difference(reference.crop(region), crop).getbbox() is not None
            assert changed == (i in (1, 2)), i
    print("edited: both target texts and original property owners verified; label and horizontal cell match")
