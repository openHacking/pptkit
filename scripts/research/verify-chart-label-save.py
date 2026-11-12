#!/usr/bin/env python3
"""Independent ZIP/XML oracle for category 1 of the two-series research fixture."""
import io
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
CHART = 'ppt/charts/chart1.xml'
BOOK = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'
SHEET = 'xl/worksheets/sheet1.xml'

def decoded(text):
    result = re.sub(r'_x([0-9A-Fa-f]{4})_', lambda m: chr(int(m.group(1), 16)), text or '')
    return result.encode('utf-16-le', 'surrogatepass').decode('utf-16-le', 'surrogatepass')

def entries(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        assert len(archive.namelist()) == len(set(archive.namelist()))
        return {name: archive.read(name) for name in archive.namelist()}

def compare_xml(before, after, paths, expected, allow_space=False):
    original, output = ET.fromstring(before), ET.fromstring(after)
    for path in paths:
        a, b = original.findall(path), output.findall(path)
        assert len(a) == len(b) == 1, path
        assert decoded(b[0].text) == expected, path
        a[0].text = b[0].text
        if allow_space and expected and (expected[0] in ' \t\r\n' or expected[-1] in ' \t\r\n'):
            space = '{http://www.w3.org/XML/1998/namespace}space'
            assert b[0].get(space) == 'preserve'
            a[0].set(space, 'preserve')
    assert ET.tostring(original) == ET.tostring(output), 'unexpected XML structure, attributes or text changes'

def verify(before, after, expected):
    old, new = entries(before), entries(after)
    assert old.keys() == new.keys()
    for name in old.keys() - {CHART, BOOK}:
        assert old[name] == new[name], name
    paths = [f'.//{C}barChart/{C}ser[{i}]/{C}cat/{C}strRef/{C}strCache/{C}pt[@idx="1"]/{C}v' for i in (1, 2)]
    compare_xml(old[CHART], new[CHART], paths, expected)
    old_book, new_book = entries(old[BOOK]), entries(new[BOOK])
    assert old_book.keys() == new_book.keys()
    for name in old_book.keys() - {SHEET}:
        assert old_book[name] == new_book[name], name
    compare_xml(old_book[SHEET], new_book[SHEET], [f'.//{S}c[@r="A3"]/{S}is/{S}t'], expected, allow_space=True)

if __name__ == '__main__':
    assert len(sys.argv) == 4, 'usage: verify-chart-label-save.py before.pptx after.pptx expected-json'
    verify(Path(sys.argv[1]).read_bytes(), Path(sys.argv[2]).read_bytes(), json.loads(sys.argv[3]))
    print('independent chart label package verification passed')
