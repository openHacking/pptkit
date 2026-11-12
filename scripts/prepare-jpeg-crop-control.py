#!/usr/bin/env python3
"""Create a physical pixel crop of the separately normalized PNG for native comparison."""
import sys
from pathlib import Path
from PIL import Image
if len(sys.argv) != 2:
    raise SystemExit("Expected sample directory.")
directory = Path(sys.argv[1])
with Image.open(directory / "browser-normalized.png") as source:
    if source.size != (200, 300):
        raise ValueError("Unexpected input dimensions")
    source.crop((0, 0, 200, 200)).save(directory / "crop-control.png")
