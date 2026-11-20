#!/usr/bin/env python3
"""Check one captured PowerPoint edit against its source, not UI provenance."""
import hashlib
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P, "a": A}
TAGS = {f"{{{P}}}{name}" for name in ("animEffect", "animRot", "animMotion")}
EXPECTED = [
    ("animEffect", {"transition": "in", "filter": "fade"}),
    ("animRot", {"by": "21600000"}),
    ("animEffect", {"transition": "out", "filter": "fade"}),
    ("animMotion", {"origin": "layout", "path": "M 0 0 L 0.25 0 E", "pathEditMode": "relative"}),
]
NAMES = ["fade-in", "spin", "fade-out", "line-path"]


def require(condition, detail):
    if not condition:
        raise ValueError(detail)


def inspect(path, edited, all_edited=False):
    with zipfile.ZipFile(path) as archive:
        for index, (tag, attrs) in enumerate(EXPECTED, 1):
            root = ET.fromstring(archive.read(f"ppt/slides/slide{index}.xml"))
            shapes = root.findall(".//p:sp", NS)
            require(len(shapes) == 1, (str(path), index, "shape count"))
            identity = shapes[0].find("p:nvSpPr/p:cNvPr", NS)
            require(identity is not None and identity.get("id") == "2" and identity.get("name") == NAMES[index - 1], (str(path), index, "shape identity"))
            text = "".join(node.text or "" for node in shapes[0].findall(".//a:t", NS))
            changed = edited and (all_edited or index == 1)
            require(text == NAMES[index - 1] + (" edited" if changed else ""), (str(path), index, "text", text))
            effects = [node for node in root.iter() if node.tag in TAGS]
            require(len(effects) == 1 and effects[0].tag == f"{{{P}}}{tag}" and effects[0].attrib == attrs, (str(path), index, "behavior"))
            times = effects[0].findall(".//p:cTn", NS)
            targets = effects[0].findall(".//p:spTgt", NS)
            require(len(times) == 1 and times[0].get("dur") == ("2000" if changed else "500"), (str(path), index, "duration"))
            require(len(targets) == 1 and targets[0].get("spid") == identity.get("id"), (str(path), index, "target"))
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def timing_tree(node):
    """Compare expanded names, attributes, text and child order; ignore indentation."""
    return (node.tag, tuple(sorted(node.attrib.items())), (node.text or "").strip(),
            (node.tail or "").strip(), tuple(timing_tree(child) for child in node))


def compare_timing(source, edited, all_edited=False):
    with zipfile.ZipFile(source) as baseline, zipfile.ZipFile(edited) as candidate:
        for index in range(1, 5):
            part = f"ppt/slides/slide{index}.xml"
            original = ET.fromstring(baseline.read(part)).find("p:timing", NS)
            updated = ET.fromstring(candidate.read(part)).find("p:timing", NS)
            require(original is not None and updated is not None, (part, "missing timing"))
            if all_edited or index == 1:
                behavior = [node for node in original.iter() if node.tag in TAGS]
                require(len(behavior) == 1, (part, "behavior count"))
                clock = behavior[0].find("p:cBhvr/p:cTn", NS)
                require(clock is not None and clock.get("dur") == "500", (part, "source duration"))
                clock.set("dur", "2000")
            require(timing_tree(original) == timing_tree(updated), (part, "unexpected timing change"))


if __name__ == "__main__":
    if len(sys.argv) not in (3, 4) or (len(sys.argv) == 4 and sys.argv[3] != "--all"):
        raise SystemExit("usage: animation-powerpoint-edit.py SOURCE EDITED [--all]")
    print("source", inspect(sys.argv[1], False))
    print("edited", inspect(sys.argv[2], True, len(sys.argv) == 4))
    compare_timing(sys.argv[1], sys.argv[2], len(sys.argv) == 4)
    print("timing trees match except permitted durations")
