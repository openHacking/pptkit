#!/usr/bin/env python3
"""Check per-object repeat and fill settings in media timing nodes."""
import json
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
with ZipFile(sys.argv[1]) as archive:
    slide = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
    found = {}
    for node in slide.findall(".//" + P + "audio/" + P + "cMediaNode"):
        target = node.find(P + "tgtEl/" + P + "spTgt")
        timing = node.find(P + "cTn")
        assert target is not None and timing is not None
        object_id = target.attrib["spid"]
        assert object_id not in found
        found[object_id] = {key: timing.attrib[key] for key in ("repeatCount", "fill") if key in timing.attrib}
    assert found == json.loads(sys.argv[2]), (found, sys.argv[2])
