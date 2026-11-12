#!/usr/bin/env python3
"""Inspect the experimental bubble worksheet, chart caches and native PDF paths."""
import hashlib
import io
import json
import platform
import posixpath
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import fitz

C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
EXPECTED = [[['1', '4'], ['2', '4'], ['1', '9']], [['7'], ['6'], ['4']]]


def require(condition, message):
    if not condition:
        raise ValueError(message)


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def package(path, source, expected=EXPECTED):
    with ZipFile(path) as archive:
        require(archive.testzip() is None, 'PPTX CRC')
        chart = ET.fromstring(archive.read('ppt/charts/chart1.xml'))
        bubble = chart.find(f'.//{C}bubbleChart')
        require(bubble is not None, 'missing native bubble chart')
        values, formulas = [], []
        for series in bubble.findall(f'{C}ser'):
            entries = [series.find(f'{C}{tag}/{C}numRef') for tag in ('xVal', 'yVal', 'bubbleSize')]
            values.append([[point.findtext(f'{C}v') for point in entry.findall(f'{C}numCache/{C}pt')] for entry in entries])
            formulas.append([entry.findtext(f'{C}f') for entry in entries])
        require(values == expected, 'X/Y/size caches differ')
        books = [name for name in archive.namelist() if name.endswith('.xlsx')]
        external = chart.find(f'{C}externalData')
        if source:
            require(len(books) == 1 and external is not None, 'source workbook missing')
            relationships = ET.fromstring(archive.read('ppt/charts/_rels/chart1.xml.rels'))
            target, = [item for item in relationships if item.get('Id') == external.get(f'{R}id')]
            require(posixpath.normpath(posixpath.join('ppt/charts', target.attrib['Target'])) == books[0], 'wrong workbook binding')
            with ZipFile(io.BytesIO(archive.read(books[0]))) as workbook:
                sheet = ET.fromstring(workbook.read('xl/worksheets/sheet1.xml'))
                cells = {cell.attrib['r']: cell.findtext(f'{S}v') for cell in sheet.iter(f'{S}c')}
                for index, (columns, data) in enumerate(zip(('ABC', 'DEF'), values)):
                    for j, (column, points) in enumerate(zip(columns, data)):
                        require(formulas[index][j] == f'Sheet1!${column}$2:${column}${len(points) + 1}', 'wrong worksheet range')
                        require([cells[f'{column}{row}'] for row in range(2, len(points) + 2)] == points, 'wrong worksheet cells')
        else:
            require(not books and external is None, 'native workbook behavior differs')
        return {'sha256': digest(path), 'values': values, 'embeddedWorkbooks': len(books), 'externalData': external is not None}


def pdf(path, ratios=(3, 2)):
    with fitz.open(path) as doc:
        require(len(doc) == 1 and 'Bubble size probe' in doc[0].get_text(), 'PDF page or title differs')
        draws = []
        for item in doc[0].get_drawings():
            rgb = tuple(round(x * 255) for x in item['fill']) if item['fill'] else None
            rect = item['rect']
            if rgb in [(68, 114, 196), (237, 125, 49)] and rect.width > 15 and abs(rect.width - rect.height) < 0.1:
                draws.append({'rgb': list(rgb), 'rectPt': list(rect), 'diameterPt': rect.width})
        draws.sort(key=lambda value: value['rectPt'][0])
        require(len(draws) == 3 and [d['rgb'] for d in draws] == [[68, 114, 196], [68, 114, 196], [237, 125, 49]], 'bubble paths differ')
        small, large, medium = [value['diameterPt'] for value in draws]
        require(abs(large / small - ratios[0]) < 0.01 and abs(medium / small - ratios[1]) < 0.01, 'area-derived diameter ratios differ')
        return {'sha256': digest(path), 'bubbles': draws}


if __name__ == '__main__':
    root = Path(sys.argv[1])
    scope = 'Public API native bubble probe' if '--public' in sys.argv[2:] else 'Internal authored native bubble probe'
    report = {'schemaVersion': 1, 'scope': f'{scope}; headless open, save and reopen only.',
              'directory': str(root), 'application': (root / 'version.txt').read_text().strip(), 'platform': platform.platform(),
              'pymupdf': fitz.VersionBind, 'source': package(root / 'bubble.pptx', True),
              'saved': package(root / 'saved/bubble.pptx', False), 'firstPdf': pdf(root / 'pdf/bubble.pdf'),
              'reopenedPdf': pdf(root / 'reopened/bubble.pdf')}
    print(json.dumps(report, indent=2))
