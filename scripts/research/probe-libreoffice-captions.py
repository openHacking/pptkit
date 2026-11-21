#!/usr/bin/env python3
"""Inspect three isolated LibreOffice resaves of the native caption fixture."""

import hashlib
import json
import subprocess
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "tests/fixtures/powerpoint-caption-loop-multitrack.pptx"
VIDEO_HASH = "dd024fc092a23ad880d94c3e3d2564056893b9647ca5f29bdb0741c787e78396"


def inspect(path):
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    containing = lambda needle: sorted(name for name, data in parts.items() if needle in data)
    return {
        "part_count": len(parts),
        "media_parts": sorted(name for name in parts if name.startswith("ppt/media/")),
        "original_video_parts": sorted(name for name, data in parts.items() if hashlib.sha256(data).hexdigest() == VIDEO_HASH),
        "vtt_parts": sorted(name for name in parts if name.lower().endswith(".vtt")),
        "track_relationship_parts": containing(b"/track"),
        "tracks_info_parts": containing(b"tracksInfo"),
        "english_cue_parts": containing(b"TWO TO FOUR"),
        "french_cue_parts": containing(b"FR QUATRE SIX"),
        "indefinite_repeat_parts": containing(b"indefinite"),
        "timing_parts": containing(b"<p:timing"),
    }


def main():
    version = subprocess.run(["soffice", "--version"], check=True, capture_output=True, text=True).stdout.strip()
    result = {"producer": version, "source_sha256": hashlib.sha256(SOURCE.read_bytes()).hexdigest(), "source": inspect(SOURCE), "rounds": []}
    with tempfile.TemporaryDirectory(prefix="pptkit-caption-probe-") as tmp:
        root = Path(tmp)
        source = SOURCE
        for index in range(1, 4):
            output = root / str(index)
            output.mkdir()
            profile = (root / ("profile-" + str(index))).as_uri()
            subprocess.run(["soffice", "-env:UserInstallation=" + profile, "--headless", "--convert-to", "pptx", "--outdir", str(output), str(source)], check=True, capture_output=True, text=True, timeout=90)
            source = output / source.name
            result["rounds"].append(inspect(source))
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
