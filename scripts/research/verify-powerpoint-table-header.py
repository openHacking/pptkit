#!/usr/bin/env python3
"""Inventory native Header Row changes in the two controlled PowerPoint saves."""
import hashlib
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
TEXT = [['Region', 'Total'], ['North', '25']]
ROW_ID = '{http://schemas.microsoft.com/office/drawing/2014/main}rowId'
COL_ID = '{http://schemas.microsoft.com/office/drawing/2014/main}colId'


def inventory(path):
    with ZipFile(path) as package:
        if package.testzip() is not None:
            raise ValueError('PPTX CRC mismatch')
        root = ET.fromstring(package.read('ppt/slides/slide1.xml'))
    table, = root.findall(f'.//{A}tbl')
    properties = table.find(f'{A}tblPr')
    rows = [[''.join(cell.itertext()) for cell in row.findall(f'{A}tc')] for row in table.findall(f'{A}tr')]
    return {
        'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
        'firstRow': properties.get('firstRow'),
        'rows': rows,
        'tableExtensionChildren': [child.tag for ext in table.iter(f'{A}ext') if ext.get('uri') for child in ext],
    }


def verify(source, off, inserted):
    states = [inventory(path) for path in (source, off, inserted)]
    if [item['firstRow'] for item in states] != ['1', None, '1']:
        raise ValueError('Expected enabled, omitted and enabled firstRow states')
    if [item['rows'] for item in states] != [TEXT, TEXT, [['', ''], *TEXT]]:
        raise ValueError('Expected insertion of just one blank first row')
    expected_extensions = [[], [COL_ID] * 2 + [ROW_ID] * 2, [COL_ID] * 2 + [ROW_ID] * 3]
    if [item['tableExtensionChildren'] for item in states] != expected_extensions:
        raise ValueError('Unexpected table extension children')
    return dict(zip(('source', 'headerOff', 'headerInserted'), states))


if __name__ == '__main__':
    if len(sys.argv) != 4:
        sys.exit('Usage: verify-powerpoint-table-header.py SOURCE.pptx OFF.pptx INSERTED.pptx')
    print(json.dumps(verify(*map(Path, sys.argv[1:])), indent=2))
