"""Check the retained first-open screenshot's two paired ink masks."""
import sys
from pathlib import Path
from PIL import Image, ImageChops

path = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/bidi-run-native/powerpoint.png')
image = Image.open(path).convert('L')
assert image.size == (1920, 1080), image.size
for y in (100, 440):
    left = image.crop((50, y, 800, y + 100)).point(lambda value: 255 if value < 100 else 0)
    right = image.crop((870, y, 1620, y + 100)).point(lambda value: 255 if value < 100 else 0)
    left_count = left.histogram()[255]; right_count = right.histogram()[255]
    difference = ImageChops.difference(left, right).histogram()[255]
    assert left_count > 1000 and right_count > 1000, (y, left_count, right_count)
    assert difference <= 50, (y, difference)
    print(f'row {y}: {left_count}/{right_count} ink pixels, {difference} differing pixels')
