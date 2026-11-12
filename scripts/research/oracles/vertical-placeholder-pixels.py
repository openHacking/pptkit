"""Verify that native inheritance matches local declarations on two marked slides."""
from pathlib import Path
import sys

from PIL import Image, ImageChops

folder = Path(sys.argv[1])
markers = []
for page in (1, 2):
    images = []
    for mode in ("inherit", "materialize"):
        with Image.open(folder / f"vertical-placeholder-{mode}-{page}.png") as image:
            assert image.size == (1920, 1080)
            images.append(image.convert("RGB"))
    for name, region in [("marker", (100, 50, 950, 155)), ("body", (100, 205, 760, 760))]:
        first, second = (im.crop(region) for im in images)
        assert ImageChops.difference(first, second).getbbox() is None, (page, name)
        bounds = first.convert("L").point(lambda v: 255 if v < 96 else 0).getbbox()
        assert bounds is not None, (page, name)
        if name == "marker":
            markers.append(first)
        else:
            left, top, right, bottom = bounds
            assert (bottom - top > right - left) == (page == 1), (page, bounds)
assert ImageChops.difference(*markers).getbbox() is not None
print("both marked pages match; inherited text is vertical and the local override is horizontal")

if "--reopens" in sys.argv[2:]:
    from posixpath import join, normpath
    from xml.etree import ElementTree as ET
    from zipfile import ZipFile

    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    for round_no in range(1, 4):
        with ZipFile(folder / f"vertical-placeholder-saved-{round_no}.pptx") as archive:
            for page, expected in ((1, None), (2, "horz")):
                slide = ET.fromstring(archive.read(f"ppt/slides/slide{page}.xml"))
                rels = ET.fromstring(archive.read(f"ppt/slides/_rels/slide{page}.xml.rels"))
                target, = [r.get("Target") for r in rels if r.get("Type", "").endswith("/slideLayout")]
                layout = ET.fromstring(archive.read(normpath(join("ppt/slides", target)).lstrip("/")))
                shape, = [s for s in slide.findall(".//p:sp", ns) if s.find("p:nvSpPr/p:nvPr/p:ph", ns) is not None]
                marker, = [s for s in slide.findall(".//p:sp", ns) if s is not shape]
                assert "".join(t.text or "" for t in marker.findall(".//a:t", ns)) == ("INHERITED VERTICAL" if page == 1 else "LOCAL HORIZONTAL")
                assert "".join(t.text or "" for t in shape.findall(".//a:t", ns)) == "中文 ABC 123"
                assert shape.find("p:txBody/a:bodyPr", ns).get("vert") == expected
                placeholder = shape.find("p:nvSpPr/p:nvPr/p:ph", ns)
                matching = []
                for candidate in layout.findall(".//p:sp", ns):
                    ph = candidate.find("p:nvSpPr/p:nvPr/p:ph", ns)
                    if ph is not None and ph.get("idx", "0") == placeholder.get("idx", "0"):
                        matching.append(candidate)
                layout_shape, = matching
                assert layout_shape.find("p:txBody/a:bodyPr", ns).get("vert") == "eaVert"
        for page in (1, 2):
            with Image.open(folder / f"vertical-placeholder-inherit-{page}.png") as initial, Image.open(
                folder / f"vertical-placeholder-reopen-{round_no}-{page}.png"
            ) as reopened:
                assert initial.size == reopened.size == (1920, 1080)
                for region in ((100, 50, 950, 155), (100, 205, 760, 760)):
                    assert ImageChops.difference(initial.convert("RGB").crop(region), reopened.convert("RGB").crop(region)).getbbox() is None
        print(f"round {round_no}: linked layout, logical text and both reopened pages verified")

if "--layout-change" in sys.argv[2:]:
    from zipfile import ZipFile

    with ZipFile(folder / "vertical-placeholder-inherit.pptx") as baseline, ZipFile(
        folder / "vertical-placeholder-layout-changed.pptx"
    ) as changed:
        assert set(baseline.namelist()) == set(changed.namelist())
        names = [name for name in baseline.namelist() if baseline.read(name) != changed.read(name)]
        name, = names
        assert name.startswith("ppt/slideLayouts/") and name.endswith(".xml")
        original = baseline.read(name)
        assert original.count(b'vert="eaVert"') == 1
        assert changed.read(name) == original.replace(b'vert="eaVert"', b'vert="vert270"')
    marker_region = (100, 50, 950, 155)
    body_region = (100, 205, 760, 760)
    for page in (1, 2):
        with Image.open(folder / f"vertical-placeholder-inherit-{page}.png") as initial, Image.open(
            folder / f"vertical-placeholder-layout-changed-{page}.png"
        ) as updated:
            assert initial.size == updated.size == (1920, 1080)
            initial, updated = initial.convert("RGB"), updated.convert("RGB")
            assert ImageChops.difference(initial.crop(marker_region), updated.crop(marker_region)).getbbox() is None
            first, second = (im.crop(body_region) for im in (initial, updated))
            if page == 1:
                a, b = (crop.convert("L").point(lambda v: 255 if v < 96 else 0).getbbox() for crop in (first, second))
                assert a and b
                assert a[0] > first.width / 2 and b[2] < second.width / 2
                assert b[3] - b[1] > b[2] - b[0]
            else:
                assert ImageChops.difference(first, second).getbbox() is None
    print("layout-only edit changes the inherited text; the local horizontal override is unchanged")
