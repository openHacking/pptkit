#!/usr/bin/env python3
"""Inspect one retained PowerPoint-produced chart/workbook without PPTKit."""
import io
import sys
import xml.etree.ElementTree as ET
import zipfile

ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart', 's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(sys.argv[1]) as presentation:
    chart = ET.fromstring(presentation.read('ppt/charts/chart1.xml'))
    books = [name for name in presentation.namelist() if name.startswith('ppt/embeddings/') and name.endswith('.xlsx')]
    assert len(books) == 1
    names = [item.text for item in chart.findall('.//c:barChart/c:ser/c:tx/c:strRef/c:strCache/c:pt/c:v', ns)]
    assert names == ['Native third', 'Second & series']
    with zipfile.ZipFile(io.BytesIO(presentation.read(books[0]))) as workbook:
        sheet = ET.fromstring(workbook.read('xl/worksheets/sheet1.xml'))
        assert sheet.find('.//s:c[@r="B1"]/s:is/s:t', ns).text == 'Native third'
        assert sheet.find('.//s:c[@r="C1"]/s:is/s:t', ns).text == 'Second & series'
print('PowerPoint chart caches and workbook series names agree')
