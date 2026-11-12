"""Compare matching PowerPoint slideshow captures, rejecting repeated pages."""
from pathlib import Path
import sys

from PIL import Image, ImageChops

folder = Path(sys.argv[1])
body = (60, 100, 1450, 390)
marker = (60, 590, 800, 760)

def dark_count(image):
    return sum(max(pixel) < 96 for pixel in image.get_flattened_data())

markers = {}
for page in (1, 2):
    images = []
    for mode in ("inherit", "materialize"):
        with Image.open(folder / f"{mode}-{page}.png") as capture:
            if capture.size != (1920, 1080):
                raise ValueError(f"{mode} page {page}: unexpected capture dimensions")
            images.append(capture.convert("RGB"))
    first, second = images
    for name, region in (("body", body), ("marker", marker)):
        a, b = first.crop(region), second.crop(region)
        if not dark_count(a) or ImageChops.difference(a, b).getbbox():
            raise ValueError(f"page {page}: blank or differing {name}")
    markers[page] = first.crop(marker)
    print(f"page {page}: matching body and marker, {dark_count(first.crop(body))} dark body pixels")
if not ImageChops.difference(markers[1], markers[2]).getbbox():
    raise ValueError("page markers repeat; capture did not distinguish pages")
