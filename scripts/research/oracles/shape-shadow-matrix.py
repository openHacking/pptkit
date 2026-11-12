#!/usr/bin/env python3
"""Inspect a retained LibreOffice shadow matrix; this does not test PowerPoint."""
import argparse
import hashlib
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile
from PIL import Image

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", type=Path)
    args = parser.parse_args()
    root = args.directory
    pptx, pdf = root / "matrix.pptx", root / "matrix.pdf"
    values = [0, .25, .5, .75, 1]
    records = []
    with zipfile.ZipFile(pptx) as archive:
        for row, shape in enumerate(values):
            for column, shadow in enumerate(values):
                index = row * 5 + column + 1
                xml = ET.fromstring(archive.read(f"ppt/slides/slide{index}.xml"))
                props = xml.find(f".//{P}sp/{P}spPr")
                fill = props.find(f"{A}solidFill/{A}srgbClr")
                effect = props.find(f"{A}effectLst/{A}outerShdw")
                color = effect.find(A + "srgbClr")
                def alpha(node):
                    element = node.find(A + "alpha")
                    return 1 if element is None else int(element.get("val")) / 100000
                if fill.get("val") != "FF0000" or color.get("val") != "000000" or alpha(fill) != shape or alpha(color) != shadow or effect.get("dist") != "1270000" or effect.get("dir") != "0":
                    raise ValueError(f"unexpected source properties on slide {index}")
                image_path = root / f"page-{index:02}.png"
                with Image.open(image_path) as image:
                    pixel = image.convert("RGB").getpixel((180, 90))
                if len(set(pixel)) != 1:
                    raise ValueError(f"non-gray shadow sample on slide {index}: {pixel}")
                records.append({"slide": index, "shapeOpacity": shape, "shadowOpacity": shadow, "pixel": list(pixel), "analyticProductGray": round(255 * (1 - shape * shadow)), "pngSha256": digest(image_path)})
    print(json.dumps({"scope": "LibreOfficeDev headless PDF export rasterized by pdftoppm at 72 dpi; no PowerPoint result", "pptxSha256": digest(pptx), "pdfSha256": digest(pdf), "samplePoint": [180, 90], "records": records}, indent=2))


if __name__ == "__main__":
    main()
