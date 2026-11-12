#!/usr/bin/env python3
"""Independent, fixture-specific chart cache and worksheet string-reference inventory."""
import io
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def decode(value):
    return re.sub(r'_x([0-9A-Fa-f]{4})_', lambda match: chr(int(match.group(1), 16)), value)

def inspect(data):
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        chart = ET.fromstring(package.read('ppt/charts/chart1.xml'))
        with zipfile.ZipFile(io.BytesIO(package.read('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'))) as workbook:
            sheet = ET.fromstring(workbook.read('xl/worksheets/sheet1.xml'))
    cells = {}
    for cell in sheet.iter(S + 'c'):
        if cell.get('t') == 'inlineStr':
            strings = cell.findall(f'{S}is/{S}t')
            assert len(strings) == 1
            cells[cell.get('r')] = decode(strings[0].text or '')
    series = chart.findall(f'.//{C}barChart/{C}ser')
    assert len(series) == 2
    names, category_sets, formulas = [], [], []
    for item in series:
        name_ref = item.find(f'{C}tx/{C}strRef')
        category_ref = item.find(f'{C}cat/{C}strRef')
        assert name_ref is not None and category_ref is not None
        def cached(ref):
            points = ref.findall(f'{C}strCache/{C}pt')
            assert [p.get('idx') for p in points] == [str(i) for i in range(len(points))]
            assert ref.find(f'{C}strCache/{C}ptCount').get('val') == str(len(points))
            return [decode(p.find(C + 'v').text or '') for p in points]
        names.append(cached(name_ref)[0]); category_sets.append(cached(category_ref))
        formulas.append((name_ref.find(C + 'f').text, category_ref.find(C + 'f').text))
    assert formulas == [('Sheet1!$B$1', 'Sheet1!$A$2:$A$4'), ('Sheet1!$C$1', 'Sheet1!$A$2:$A$4')]
    assert names == [cells['B1'], cells['C1']] == ['First _x005F_', 'Second & series']
    categories = [cells['A2'], cells['A3'], cells['A4']]
    assert category_sets == [categories, categories] == [[' Q1 ', 'literal _x0041_', '中文 & two']] * 2
    return {'series': names, 'categories': categories, 'categoryCacheCopies': len(category_sets), 'sharedCategoryRange': formulas[0][1]}

if __name__ == '__main__':
    assert len(sys.argv) == 2, 'usage: verify-chart-label-source.py source.pptx'
    print(inspect(Path(sys.argv[1]).read_bytes()))
