#!/usr/bin/env python3
"""Check observed Chromium sample pixels; screenshots do not prove timing or provenance."""
import argparse
import hashlib
import json
from pathlib import Path
from PIL import Image


def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()


def inspect(root):
    modes = {}
    for mode in ("inline", "svg-image", "direct-image"):
        times = json.loads((root / f"{mode}-times.json").read_text())
        if len(times) != 10 or any(a >= b for a, b in zip(times, times[1:])): raise ValueError("sample times")
        samples = []
        for i, elapsed in enumerate(times):
            path = root / f"{mode}-{i}.png"
            with Image.open(path) as image:
                if image.size != (960, 540): raise ValueError("screenshot dimensions")
                pixel = image.convert("RGB").getpixel((400, 250))
            samples.append({"elapsedBeforeScreenshotMs": elapsed, "pixel": list(pixel), "sha256": digest(path)})
        colors = {tuple(s["pixel"]) for s in samples}
        if (255, 0, 0) not in colors or (127, 0, 128) not in colors: raise ValueError(f"missing observed frame colors: {mode}")
        modes[mode] = samples
    return {"scope": "one APNG in Chromium; screenshot samples do not prove timing, UI provenance or other viewers", "svgSha256": digest(root / "preview.svg"), "samplePoint": [400,250], "modes": modes}


def main():
    parser=argparse.ArgumentParser();parser.add_argument("directory",type=Path);args=parser.parse_args();print(json.dumps(inspect(args.directory),indent=2))


if __name__ == "__main__": main()
