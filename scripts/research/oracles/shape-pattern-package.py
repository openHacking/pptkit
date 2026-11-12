#!/usr/bin/env python3
"""Check the deliberately narrow three-slide native pattern probe, not rendering."""
import argparse
import io
import xml.etree.ElementTree as ET
import zipfile

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def require(condition, message):
    if not condition:
        raise ValueError(message)


def verify(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
        size = presentation.find(P + "sldSz")
        require(size is not None and (size.get("cx"), size.get("cy")) == ("12192000", "6858000"), "slide size")
        rels = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels.findall(REL + "Relationship")}
        slides = presentation.findall(f"{P}sldIdLst/{P}sldId")
        require(len(slides) == 3, "slide count")
        for index, (slide, preset) in enumerate(zip(slides, ("horz", "vert", "cross"))):
            target = targets.get(slide.get(R + "id"))
            require(target == f"slides/slide{index + 1}.xml", f"slide {index + 1} order")
            root = ET.fromstring(archive.read("ppt/" + target))
            shapes = root.findall(f".//{P}spTree/{P}sp")
            require(len(shapes) == 1, f"slide {index + 1} shape count")
            shape = shapes[0]
            identity = shape.find(f"{P}nvSpPr/{P}cNvPr")
            require(identity is not None and identity.get("name") == f"shape-{index}", "shape identity")
            props = shape.find(P + "spPr")
            off, ext = props.find(f"{A}xfrm/{A}off"), props.find(f"{A}xfrm/{A}ext")
            require(off is not None and (off.get("x"), off.get("y")) == ("2286000", "1524000"), "position")
            require(ext is not None and (ext.get("cx"), ext.get("cy")) == ("7620000", "3810000"), "size")
            geom = props.find(A + "prstGeom")
            require(geom is not None and geom.get("prst") == "rect", "geometry")
            fill = props.find(A + "pattFill")
            require(fill is not None and fill.get("prst") == preset, "pattern order")
            fg, bg = fill.find(f"{A}fgClr/{A}srgbClr"), fill.find(f"{A}bgClr/{A}srgbClr")
            require(fg is not None and fg.get("val") == "FF0000", "foreground")
            alpha = fg.find(A + "alpha")
            require(alpha is not None and alpha.get("val") == "50000", "foreground alpha")
            require(bg is not None and bg.get("val") == "0000FF" and bg.find(A + "alpha") is None, "background")
            require(props.find(f"{A}ln/{A}noFill") is not None, "stroke")


def mutate(data, part, before, after):
    source = zipfile.ZipFile(io.BytesIO(data))
    result = io.BytesIO()
    with source, zipfile.ZipFile(result, "w") as target:
        for info in source.infolist():
            content = source.read(info.filename)
            if info.filename == part:
                require(before in content, "negative control did not alter fixture")
                content = content.replace(before, after, 1)
            target.writestr(info, content)
    return result.getvalue()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx")
    args = parser.parse_args()
    with open(args.pptx, "rb") as source:
        data = source.read()
    verify(data)
    for before, after in ((b'prst="horz"', b'prst="vert"'), (b'val="50000"', b'val="100000"'), (b'x="2286000"', b'x="2286001"'), (b'val="0000FF"', b'val="00FF00"')):
        altered = mutate(data, "ppt/slides/slide1.xml", before, after)
        try:
            verify(altered)
        except ValueError:
            continue
        raise AssertionError(f"negative control accepted: {before!r}")
    print("Pattern probe package verified; four mutation controls rejected. Native appearance is not assessed.")


if __name__ == "__main__":
    main()
