#!/usr/bin/env python3
"""Inspect already captured PowerPoint saves; this does not operate PowerPoint."""
import hashlib
import json
import platform
import posixpath
import re
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

BASE = Path(".tmp/cross-document-copy-native")
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R = "http://schemas.openxmlformats.org/package/2006/relationships"

def digest(data):
    return hashlib.sha256(data).hexdigest()

def inspect(file):
    with zipfile.ZipFile(file) as z:
        assert sorted(n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)) == ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]
        assert sorted(n for n in z.namelist() if re.fullmatch(r"ppt/notesSlides/notesSlide\d+\.xml", n)) == ["ppt/notesSlides/notesSlide1.xml", "ppt/notesSlides/notesSlide2.xml"]
        def text(part):
            return "".join(x.text or "" for x in ET.fromstring(z.read(part)).iter(f"{{{A}}}t"))
        def rels(part):
            folder, name = posixpath.split(part)
            relpart = f"{folder}/_rels/{name}.rels"
            result = {}
            for rel in ET.fromstring(z.read(relpart)):
                target = rel.get("Target")
                result[rel.get("Id")] = {"type": rel.get("Type", "").split("/")[-1],
                    "target": posixpath.normpath(posixpath.join(folder, target)),
                    "mode": rel.get("TargetMode", "Internal")}
            return result
        slides = [f"ppt/slides/slide{i}.xml" for i in (1, 2)]
        notes = [f"ppt/notesSlides/notesSlide{i}.xml" for i in (1, 2)]
        assert [text(x) for x in slides] == ["Control page", "Copied page"]
        assert [text(x) for x in notes] == ["Control page notes", "Copied page notes"]
        pairs = []
        for slide, note in zip(slides, notes):
            assert any(r["type"] == "notesSlide" and r["target"] == note and r["mode"] == "Internal" for r in rels(slide).values())
            assert any(r["type"] == "slide" and r["target"] == slide and r["mode"] == "Internal" for r in rels(note).values())
            pairs.append({"slide": slide, "notes": note, "notesText": text(note)})
        tree = ET.fromstring(z.read(slides[1])).find(f".//{{{P}}}spTree")
        objects = []
        for element in tree:
            kind = element.tag.split("}")[-1]
            if kind not in ("sp", "pic", "cxnSp"):
                continue
            props = element.find(f".//{{{P}}}cNvPr")
            transform = element.find(f"{{{P}}}spPr/{{{A}}}xfrm")
            off = transform.find(f"{{{A}}}off") if transform is not None else None
            ext = transform.find(f"{{{A}}}ext") if transform is not None else None
            click = props.find(f"{{{A}}}hlinkClick") if props is not None else None
            objects.append({"kind": kind, "name": props.get("name"),
                "offset": dict(off.attrib) if off is not None else None,
                "extent": dict(ext.attrib) if ext is not None else None,
                "customGeometry": element.find(f".//{{{A}}}custGeom") is not None,
                "clickId": click.get(f"{{{O}}}id") if click is not None else None})
        byname = {x["name"]: x for x in objects}
        assert set(byname) == {"element-1", "element-2", "element-3", "element-4"}
        assert byname["element-2"]["kind"] == "pic" and byname["element-4"]["customGeometry"]
        click = rels(slides[1])[byname["element-1"]["clickId"]]
        assert click == {"type": "slide", "target": slides[1], "mode": "Internal"}
        media = [{"name": n, "sha256": digest(z.read(n))} for n in z.namelist() if n.startswith("ppt/media/") and not n.endswith("/")]
        assert len(media) == 1 and media[0]["sha256"] == "45ca74602a2f7790a8eb2fad662cc73ee8ab42e78d5abfd0c66aa4d44725c9c6"
        return {"file": str(file.resolve()), "sha256": digest(file.read_bytes()), "slideText": [text(x) for x in slides],
                "pageNotesPairs": pairs, "copiedPageObjects": objects, "selfLink": click, "media": media}

files = [BASE / "input.pptx"] + [BASE / f"powerpoint-round{i}.pptx" for i in (1, 2, 3)]
rounds = [inspect(file) for file in files]
first = {x["name"]: x for x in rounds[0]["copiedPageObjects"]}
last = {x["name"]: x for x in rounds[-1]["copiedPageObjects"]}
for item in rounds[1:]:
    current = {x["name"]: x for x in item["copiedPageObjects"]}
    for name in first:
        assert current[name]["offset"] == first[name]["offset"]
        assert current[name]["extent"] == first[name]["extent"]
assert digest((BASE / "powerpoint-work.pptx").read_bytes()) == digest(files[0].read_bytes())
deltas = {}
for name in first:
    deltas[name] = {}
    for part in ("offset", "extent"):
        if first[name][part] is not None and last[name][part] is not None:
            deltas[name][part] = {k: int(v) - int(first[name][part][k]) for k, v in last[name][part].items()}
shots = []
for prefix in ("first", "round1", "round2", "round3"):
    for slide in (1, 2):
        name = f"{prefix}-slide{slide}" + ("-notes" if prefix == "first" else "") + ".png"
        file = BASE / name
        shots.append({"file": str(file.resolve()), "sha256": digest(file.read_bytes())})
version = subprocess.check_output(["defaults", "read", "/Applications/Microsoft PowerPoint.app/Contents/Info", "CFBundleShortVersionString"], text=True).strip()
evidence = {"schemaVersion": 1, "scope": "Two-page copied-slide fixture; PowerPoint three successive Save As and reopen observations. Package inventory is independent. No claim of link activation, editable connectors, or full visual fidelity.",
    "environment": {"application": "Microsoft PowerPoint " + version, "system": platform.platform(), "architecture": platform.machine()},
    "observations": "Original and each saved file were opened in PowerPoint; each page and its notes were inspected visually. No repair prompt was observed. Screenshots are local temporary artifacts.",
    "independent": {"oracle": "Python zipfile and ElementTree", "rounds": rounds, "finalGeometryDeltasEmu": deltas}, "screenshots": shots}
Path("docs/testing/cross-copy-powerpoint-evidence.json").write_text(json.dumps(evidence, indent=2) + "\n")
print(json.dumps({"application": version, "files": len(rounds), "screenshots": len(shots), "finalGeometryDeltasEmu": deltas}, indent=2))
