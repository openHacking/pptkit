#!/usr/bin/env python3
"""Create the local red/blue playback probe; Pillow is a research dependency."""
from pathlib import Path
import sys
from PIL import Image

if len(sys.argv) != 2:
    raise SystemExit('Usage: python3 scripts/prepare-gif-native-source.py <new-directory>')
directory = Path(sys.argv[1]); directory.mkdir(parents=True, exist_ok=False)
frames = [Image.new('RGB', (40, 40), color) for color in ('red', 'blue')]
frames[0].save(directory / 'motion.gif', save_all=True, append_images=frames[1:], duration=[1500, 1500], loop=0, disposal=2)
