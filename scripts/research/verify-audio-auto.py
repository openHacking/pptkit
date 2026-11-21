#!/usr/bin/env python3
"""Independently inspect main media sequence, offsets and targets."""
import json
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
with ZipFile(sys.argv[1]) as z:
    root = ET.fromstring(z.read("ppt/slides/slide1.xml"))
    timing = root.find(P + "timing")
    ids = [n.get("id") for n in timing.iter(P + "cTn")]
    assert len(ids) == len(set(ids))
    seqs = [n for n in timing.iter(P + "seq") if n.find(P + "cTn").get("nodeType") == "mainSeq"]
    assert len(seqs) == 1
    main = seqs[0].find(P + "cTn")
    found = []
    for group_index, par in enumerate(main.find(P + "childTnLst")):
        group = par.find(P + "cTn")
        conditions = group.find(P + "stCondLst")
        begin = [[c.get("evt"), c.get("delay"), c.find(P + "tn").get("val") if c.find(P + "tn") is not None else None] for c in conditions]
        effects = []
        for wrapper in group.find(P + "childTnLst"):
            outer = wrapper.find(P + "cTn")
            effect = next(n for n in outer.iter(P + "cTn") if n.get("presetClass") == "mediacall")
            cmd = effect.find(".//" + P + "cmd")
            target = cmd.find(P + "cBhvr/" + P + "tgtEl/" + P + "spTgt").get("spid")
            duration = cmd.find(P + "cBhvr/" + P + "cTn").get("dur")
            effects.append({"target": target, "type": effect.get("nodeType"), "delay": outer.find(P + "stCondLst/" + P + "cond").get("delay"), "duration": duration})
        found.append({"begin": begin, "effects": effects})
    assert found == json.loads(sys.argv[2]), (found, sys.argv[2])
