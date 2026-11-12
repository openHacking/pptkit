#!/usr/bin/env python3
"""Independent narrow oracle for the repository's named field/chart-save sample."""
import io
import sys
import zipfile
from xml.etree import ElementTree as ET

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
PRESENTATION = 'ppt/presentation.xml'
SLIDE = 'ppt/slides/slide1.xml'
CHART = 'ppt/charts/chart1.xml'
BOOK = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'
SHEET = 'xl/worksheets/sheet1.xml'


def parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Duplicate ZIP members')
        return {name: archive.read(name) for name in names}


def same_xml_after(before, after, edit):
    original, candidate = ET.fromstring(before), ET.fromstring(after)
    edit(original)
    if ET.tostring(original) != ET.tostring(candidate):
        raise ValueError('Unexpected XML change')


def verify(source, candidate):
    before, after = parts(source), parts(candidate)
    if before.keys() != after.keys():
        raise ValueError('Package inventory changed')
    changed = {name for name in before if before[name] != after[name]}
    if changed != {PRESENTATION, SLIDE, CHART, BOOK}:
        raise ValueError(f'Unexpected outer changes: {changed}')
    def number(root):
        if root.tag != P + 'presentation' or 'firstSlideNum' in root.attrib:
            raise ValueError('Unexpected source presentation')
        root.set('firstSlideNum', '17')
    def title(root):
        fields = list(root.iter(A + 'fld'))
        if len(fields) != 1 or fields[0].get('id', '').upper() != '{11111111-2222-3333-4444-555555555555}' or fields[0].get('type') != 'slidenum' or fields[0].findtext(A + 't') != 'STALE':
            raise ValueError('Unexpected source field')
        matches = [node for node in root.iter(A + 't') if node.text == 'Page: ']
        if len(matches) != 1:
            raise ValueError('Ambiguous title')
        matches[0].text = 'Edited page: '
    def chart(root):
        series = list(root.iter(C + 'ser'))
        if not series:
            raise ValueError('Missing chart series')
        val = series[0].find(C + 'val')
        points = [] if val is None else [p for p in val.iter(C + 'pt') if p.get('idx') == '1']
        if len(points) != 1 or points[0].findtext(C + 'v') != '5':
            raise ValueError('Unexpected chart value')
        points[0].find(C + 'v').text = '7'
    def sheet(root):
        cells = [node for node in root.iter(S + 'c') if node.get('r') == 'B3']
        if len(cells) != 1 or cells[0].findtext(S + 'v') != '5':
            raise ValueError('Unexpected worksheet value')
        cells[0].find(S + 'v').text = '7'
    for name, edit in [(PRESENTATION, number), (SLIDE, title), (CHART, chart)]:
        same_xml_after(before[name], after[name], edit)
    old_book, new_book = parts(before[BOOK]), parts(after[BOOK])
    if old_book.keys() != new_book.keys() or {n for n in old_book if old_book[n] != new_book[n]} != {SHEET}:
        raise ValueError('Unexpected workbook changes')
    same_xml_after(old_book[SHEET], new_book[SHEET], sheet)
    return {'outer_changed': sorted(changed), 'workbook_changed': [SHEET], 'first_slide_number': 17, 'chart_value': 7, 'cell': 'B3', 'field_type': 'slidenum', 'field_cache': 'STALE'}


if __name__ == '__main__':
    import json
    with open(sys.argv[1], 'rb') as f:
        source = f.read()
    with open(sys.argv[2], 'rb') as f:
        candidate = f.read()
    print(json.dumps(verify(source, candidate), sort_keys=True))
