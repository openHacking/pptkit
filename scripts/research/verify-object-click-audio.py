#!/usr/bin/env python3
"""Check ordered main effects and independently targeted media interactions."""
import json
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
with ZipFile(sys.argv[1]) as package:
    root = ET.fromstring(package.read('ppt/slides/slide1.xml'))
expected = json.loads(sys.argv[2])
timing = root.find('p:timing', ns)
ids = [n.get('id') for n in timing.findall('.//p:cTn', ns)]
assert len(ids) == len(set(ids))
sequences = timing.findall('.//p:seq', ns)
main = [s for s in sequences if s.find('p:cTn', ns).get('nodeType') == 'mainSeq']
assert len(main) == (1 if expected['main'] else 0)
def command_target(cmd):
    return cmd.find('p:cBhvr/p:tgtEl/p:spTgt', ns).get('spid')
assert [(c.get('cmd'), command_target(c)) for c in main[0].findall('.//p:cmd', ns)] == [('playFrom(0.0)', x) for x in expected['main']] if main else not expected['main']
interactions = [s for s in sequences if s.find('p:cTn', ns).get('nodeType') == 'interactiveSeq']
seen = []
for sequence in interactions:
    start = sequence.find('p:cTn/p:stCondLst/p:cond', ns)
    next_cond = sequence.find('p:nextCondLst/p:cond', ns)
    target = start.find('p:tgtEl/p:spTgt', ns).get('spid')
    assert start.get('evt') == next_cond.get('evt') == 'onClick'
    assert next_cond.find('p:tgtEl/p:spTgt', ns).get('spid') == target
    cmds = sequence.findall('.//p:cmd', ns)
    assert len(cmds) == 1 and command_target(cmds[0]) == target
    seen.append([target, cmds[0].get('cmd')])
assert seen == [[x, 'togglePause'] for x in expected['videos']] + [[x, 'playFrom(0.0)'] for x in expected['objects']]
for kind in ('audio', 'video'):
    actual = [x.find('p:cMediaNode/p:tgtEl/p:spTgt', ns).get('spid') for x in timing.findall('.//p:' + kind, ns)]
    assert actual == expected[kind]
print('Verified object-click media timing')
