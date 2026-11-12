"""Check the observed geometry of the seven native slideshow captures."""
from pathlib import Path
import sys
from PIL import Image, ImageChops

folder = Path(sys.argv[1])
boxes = [(100, 155, 770, 850), (880, 155, 1270, 850), (1270, 155, 1690, 850)]
modes = ["horz", "vert", "vert270", "eaVert", "mongolianVert", "wordArtVert", "wordArtVertRtl"]
labels = set()
horizontal_cell = None
for page, mode in enumerate(modes, 1):
    with Image.open(folder / f"mode-{page}.png") as capture:
        if capture.size != (1920, 1080):
            raise ValueError(f"{mode}: unexpected capture dimensions")
        im = capture.convert("L")
    label = im.crop((100, 55, 530, 145)).point(lambda v: 255 if v < 96 else 0)
    if not label.getbbox() or label.tobytes() in labels:
        raise ValueError(f"{mode}: blank or repeated label")
    labels.add(label.tobytes())
    crops = [im.crop(box) for box in boxes]
    for j, crop in enumerate(crops):
        bounds = crop.point(lambda v: 255 if v < 96 else 0).getbbox()
        if bounds is None:
            raise ValueError(f"{mode}: blank text in region {j}")
        left, top, right, bottom = bounds
        vertical = j < 2 and page != 1
        if (bottom - top > right - left) != vertical:
            raise ValueError(f"{mode}: wrong axis in region {j}")
        if vertical and ((left > crop.width / 2) != (page in {2, 4, 7})):
            raise ValueError(f"{mode}: wrong side in region {j}")
    if horizontal_cell is None:
        horizontal_cell = crops[2]
    elif ImageChops.difference(horizontal_cell, crops[2]).getbbox():
        raise ValueError(f"{mode}: bodyPr-only cell differs from the horizontal control")
    print(f"{mode}: distinct label, expected axes and anchors")
