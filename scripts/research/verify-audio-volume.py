#!/usr/bin/env python3
"""Compare target-specific native audio volume attributes in a PPTX."""
import json
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
with ZipFile(sys.argv[1]) as archive:
    slide = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
    nodes = slide.findall(".//" + P + "audio/" + P + "cMediaNode")
    found = {}
    for node in nodes:
        target = node.find(P + "tgtEl/" + P + "spTgt")
        assert target is not None
        object_id = target.attrib["spid"]
        assert object_id not in found
        found[object_id] = {key: node.attrib[key] for key in ("vol", "mute") if key in node.attrib}
    assert found == json.loads(sys.argv[2]), (found, sys.argv[2])
