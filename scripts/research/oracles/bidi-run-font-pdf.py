#!/usr/bin/env python3
"""Check observed Arabic and digit PDF font traces in both native controls."""
import sys
import fitz

assert len(sys.argv) == 3
for path, arabic_font in ((sys.argv[1], 'ArialMT'), (sys.argv[2], 'TimesNewRomanPSMT')):
    document = fitz.open(path)
    assert len(document) == 1
    spans = document[0].get_texttrace()
    digits = [s for s in spans if ''.join(chr(c[0]) for c in s['chars']) == '123']
    arabic = [s for s in spans if any(0xFB50 <= c[0] <= 0xFEFF for c in s['chars'])]
    for label, selected, expected in (('digits', digits, 'ArialMT'), ('arabic', arabic, arabic_font)):
        assert len(selected) == 4, (path, label, len(selected))
        assert all(s['font'] == expected and len(s['chars']) > 0 and all(c[1] != 0 for c in s['chars']) for s in selected), (path, label)
        quadrants = {(s['bbox'][0] >= 400, s['bbox'][1] >= 150) for s in selected}
        assert len(quadrants) == 4, (path, label, quadrants)
    print(path, 'four Arabic and four digit traces verified')
