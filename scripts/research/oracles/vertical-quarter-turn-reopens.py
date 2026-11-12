"""Check first-open and save/reopen captures for one three-mode profile."""
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
profile = sys.argv[2] if len(sys.argv) > 2 else "quarter"
profiles = {
    "quarter": ("quarter-turn", ("horz", "vert", "vert270"), 0),
    "remaining": ("stacked-mongolian", ("mongolianVert", "wordArtVert", "wordArtVertRtl"), 4),
}
package, modes, source_offset = profiles[profile]
body = (100, 155, 770, 850)
marker = (100, 55, 530, 145)

def crop(prefix, page, box):
    with Image.open(folder / f"{prefix}-{page}.png") as im:
        assert im.size == (1920, 1080), (prefix, page, im.size)
        return im.convert("RGB").crop(box)

for round_no in range(4):
    suffix = "pre-save" if round_no == 0 else f"saved-{round_no}"
    with ZipFile(folder / f"{package}-{suffix}.pptx") as archive:
        for page, mode in enumerate(modes, 1):
            slide = ET.fromstring(archive.read(f"ppt/slides/slide{page}.xml"))
            label, shape = slide.findall(".//p:sp/p:txBody", ns)
            assert "".join(label.itertext()).strip() == "MODE: " + mode, (suffix, page)
            assert shape.find("a:bodyPr", ns).get("vert") == mode, (suffix, page)
            assert "".join(t.text or "" for t in shape.findall(".//a:t", ns)) == "日本語（ABC 123）", (suffix, page)
            if profile == "remaining":
                vertical, horizontal = slide.findall(".//a:tc", ns)
                assert vertical.find("a:tcPr", ns).get("vert") == mode, (suffix, page)
                assert horizontal.find("a:tcPr", ns).get("vert") is None, (suffix, page)
                for cell in (vertical, horizontal):
                    assert cell.find("a:txBody/a:bodyPr", ns).get("vert") is None, (suffix, page)
                    assert "".join(t.text or "" for t in cell.findall(".//a:t", ns)) == "日本語（ABC 123）", (suffix, page)
    print(suffix, "verified")
regions = (body, marker) if profile == "quarter" else (body, marker, (880, 155, 1270, 850), (1270, 155, 1690, 850))
for page in range(1, 4):
    for box in regions:
        reference = crop("mode", page + source_offset, box)
        assert min(reference.convert("L").getextrema()) < 96, (page, box)
        for prefix in (f"{profile}-first", *(f"{profile}-reopen-{n}" for n in range(1, 4))):
            assert ImageChops.difference(reference, crop(prefix, page, box)).getbbox() is None, (prefix, page, box)
    print("page", page, "matched the original in all four comparisons")
