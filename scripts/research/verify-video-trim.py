#!/usr/bin/env python3
"""Check source-relative trims and their command durations by shape target."""
import json
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
P14 = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
with ZipFile(sys.argv[1]) as archive:
    root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
    found = {}
    for pic in root.findall(".//" + P + "spTree/" + P + "pic"):
        obj = pic.find(P + "nvPicPr/" + P + "cNvPr")
        media = pic.find(".//" + P14 + "media")
        if obj is None or media is None or pic.find(".//" + A + "videoFile") is None:
            continue
        trim = media.find(P14 + "trim")
        found[obj.attrib["id"]] = {"trim": None if trim is None else {"st": trim.get("st", "0"), "end": trim.get("end", "0")}}
    for cmd in root.findall(".//" + P + "cmd"):
        if cmd.get("cmd") != "playFrom(0.0)":
            continue
        target = cmd.find(P + "cBhvr/" + P + "tgtEl/" + P + "spTgt")
        time = cmd.find(P + "cBhvr/" + P + "cTn")
        assert target is not None and time is not None
        ident = target.attrib["spid"]
        if ident in found:
            assert "duration" not in found[ident]
            found[ident]["duration"] = time.attrib["dur"]
    assert found == json.loads(sys.argv[2]), (found, sys.argv[2])
