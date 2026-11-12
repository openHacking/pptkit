"""Independently inspect the transformed native shape text orientation."""
import json
import math
import sys
from pathlib import Path
from xml.etree import ElementTree

import fitz

root = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/native-mixed-shape-svg-probe')
svg = ElementTree.parse(root / 'native.svg')
ns = {'s': 'http://www.w3.org/2000/svg'}
group = next(e for e in svg.findall('.//s:g', ns) if e.get('data-pptkit-element-id') == 'mixed-shape')
assert 'rotate(8)' in group.get('transform', '')
assert group.find('s:text', ns) is not None
doc = fitz.open(root / 'native.svg')
try:
    lines = [line for block in doc[0].get_text('dict')['blocks'] if 'lines' in block for line in block['lines']]
    assert len(lines) == 1
    angle = math.degrees(math.atan2(lines[0]['dir'][1], lines[0]['dir'][0]))
    assert abs(angle - 8) < 0.01
    assert ''.join(span['text'] for span in lines[0]['spans']) == 'Alpha Beta'
finally:
    doc.close()
print(json.dumps({'observedAngleDegrees': angle, 'text': 'Alpha Beta'}))
