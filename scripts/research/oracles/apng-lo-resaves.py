#!/usr/bin/env python3
"""Inspect one-image APNG source and three headless LibreOffice resaves."""
import argparse
import hashlib
import json
import posixpath
from pathlib import Path
import struct
import xml.etree.ElementTree as ET
import zipfile
import zlib

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def sha(data):
    return hashlib.sha256(data).hexdigest()


def animation(data):
    if data[:8] != b"\x89PNG\r\n\x1a\n": raise ValueError("PNG signature")
    offset, controls, ended = 8, [], False
    while offset < len(data):
        size = struct.unpack_from(">I", data, offset)[0]
        kind, payload = data[offset+4:offset+8], data[offset+8:offset+8+size]
        if offset+size+12 > len(data) or zlib.crc32(kind+payload) != struct.unpack_from(">I", data, offset+8+size)[0]: raise ValueError("chunk or CRC")
        if kind == b"acTL": controls.append(["acTL", *struct.unpack(">II", payload)])
        if kind == b"fcTL": controls.append(["fcTL", *struct.unpack(">IIIIIHHBB", payload)])
        if kind == b"fdAT": controls.append(["fdAT", struct.unpack_from(">I", payload)[0]])
        offset += size+12
        if kind == b"IEND": ended = True; break
    if not ended or offset != len(data): raise ValueError("IEND or trailing bytes")
    return controls


def inspect(path, expected):
    raw = path.read_bytes()
    with zipfile.ZipFile(path) as archive:
        slide = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
        pictures = slide.findall(f".//{P}pic")
        if len(pictures) != 1: raise ValueError("picture count")
        picture = pictures[0]
        blip = picture.find(f"{P}blipFill/{A}blip")
        rels = ET.fromstring(archive.read("ppt/slides/_rels/slide1.xml.rels"))
        matches = [r for r in rels.findall(REL+"Relationship") if r.get("Id") == blip.get(R+"embed")]
        if len(matches) != 1 or not matches[0].get("Type", "").endswith("/image") or matches[0].get("TargetMode") == "External": raise ValueError("image relationship")
        media = posixpath.normpath(posixpath.join("ppt/slides", matches[0].get("Target")))
        image = archive.read(media)
        if image != expected: raise ValueError("image bytes changed")
        transform = picture.find(f"{P}spPr/{A}xfrm")
        off, ext = transform.find(A+"off"), transform.find(A+"ext")
        return {"fileSha256": sha(raw), "media": media, "mediaSha256": sha(image), "boxEmu": [int(off.get("x")),int(off.get("y")),int(ext.get("cx")),int(ext.get("cy"))]}


def main():
    parser=argparse.ArgumentParser();parser.add_argument("directory",type=Path);args=parser.parse_args();root=args.directory
    expected=Path("tests/fixtures/two-frame-apng.png").read_bytes();controls=animation(expected)
    if not any(item[0]=="acTL" and item[1:]==[2,2] for item in controls) or len([x for x in controls if x[0]=="fcTL"])!=2: raise ValueError("fixture animation controls")
    outputs={name:inspect(root/name/"source.pptx",expected) for name in ("source","round1","round2","round3")}
    print(json.dumps({"scope":"one PPTKit APNG fixture, LibreOfficeDev headless resave; no playback or PowerPoint evidence", "animationControls":controls,"outputs":outputs},indent=2))


if __name__ == "__main__": main()
