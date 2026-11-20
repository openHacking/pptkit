#!/usr/bin/env python3
"""Exercise the optional native-artifact oracle, including Python optimized mode."""
from pathlib import Path
import subprocess
import sys
import tempfile
import zipfile

ORACLE = Path(__file__).with_name("animation-powerpoint-edit.py")
CASES = (
    (1, b'dur="2000"', b'dur="500"', "duration"),
    (2, b'spid="2"', b'spid="999"', "target"),
    (3, b'fade-out edited', b'wrong text', "text"),
    (4, b'pathEditMode="relative"', b'pathEditMode="fixed"', "behavior"),
    (1, b'evt="onNext"', b'evt="onPrev"', "unexpected timing change"),
    (2, b'restart="never"', b'restart="always"', "unexpected timing change"),
)


def run(source, candidate, optimized):
    options = ["-O"] if optimized else []
    return subprocess.run([sys.executable, *options, str(ORACLE), str(source), str(candidate), "--all"],
                          capture_output=True, text=True, check=False)


def mutated(source, destination, slide, old, new):
    part = f"ppt/slides/slide{slide}.xml"
    changed = False
    with zipfile.ZipFile(source) as reader, zipfile.ZipFile(destination, "w") as writer:
        for item in reader.infolist():
            data = reader.read(item.filename)
            if item.filename == part:
                if old not in data:
                    raise ValueError(f"mutation anchor absent in {part}: {old!r}")
                data = data.replace(old, new, 1)
                changed = True
            writer.writestr(item, data)
    if not changed:
        raise ValueError(f"part absent: {part}")


def main():
    if len(sys.argv) != 3:
        raise SystemExit("usage: test-animation-powerpoint-edit.py SOURCE FOUR_EDIT_FILE")
    source, candidate = map(Path, sys.argv[1:])
    with tempfile.TemporaryDirectory() as directory:
        for optimized in (False, True):
            valid = run(source, candidate, optimized)
            if valid.returncode != 0:
                raise RuntimeError(f"valid artifact rejected, optimized={optimized}: {valid.stderr}")
            for index, (slide, old, new, expected_reason) in enumerate(CASES):
                altered = Path(directory) / f"{index}.pptx"
                mutated(candidate, altered, slide, old, new)
                result = run(source, altered, optimized)
                if result.returncode == 0 or "ValueError" not in result.stderr or expected_reason not in result.stderr:
                    raise RuntimeError(f"mutation {index} not rejected for {expected_reason}, optimized={optimized}: {result.stderr}")
            print(f"optimized={optimized}: valid accepted; {len(CASES)} mutations rejected")


if __name__ == "__main__":
    main()
