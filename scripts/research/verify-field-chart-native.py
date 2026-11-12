#!/usr/bin/env python3
"""Check the retained PowerPoint field/chart rounds against their expected page order."""
import hashlib
import importlib.util
import io
import json
import subprocess
import sys
import zipfile
from xml.etree import ElementTree as ET

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def verify(filename, position):
    raw = open(filename, 'rb').read()
    report = json.loads(subprocess.check_output([sys.executable, 'scripts/research/chart-value-save-oracle.py', 'inspect', filename]))
    charts = report['charts']
    if len(charts) != 4 or [c['slideIndex'] for c in charts] != [0, 1, 2, 3]:
        raise ValueError('Unexpected chart or slide inventory')
    target = charts[position]
    if target['family'] != 'bar' or target['series'][0]['values'] != [2, 7, 8] or target['series'][1]['values'] != [3, 4, 6]:
        raise ValueError('Target cache changed')
    if any(chart['series'][0]['values'] != [2, 5, 8] for i, chart in enumerate(charts) if i != position):
        raise ValueError('Control chart values changed')
    with zipfile.ZipFile(io.BytesIO(raw)) as deck:
        presentation = ET.fromstring(deck.read('ppt/presentation.xml'))
        if presentation.get('firstSlideNum') != '17':
            raise ValueError('Starting number changed')
        spec = importlib.util.spec_from_file_location('chart_oracle', 'scripts/research/chart-value-save-oracle.py')
        oracle = importlib.util.module_from_spec(spec); spec.loader.exec_module(oracle)
        package = oracle.package(raw); main = 'ppt/presentation.xml'; relations = oracle.rels(package, main)
        slide_ids = presentation.findall(f'{P}sldIdLst/{P}sldId')
        if len(slide_ids) != 4:
            raise ValueError('Unexpected slide count')
        expected_titles = {'line': 'Line values', 'pie': 'Pie values', 'bar': 'Unchanged control'}
        for index, slide_id in enumerate(slide_ids):
            slide_part = oracle.target(main, relations.get(slide_id.get('{' + oracle.R + '}id')))
            slide = ET.fromstring(deck.read(slide_part))
            expected = 'Bar values' if index == position else expected_titles[charts[index]['family']]
            if [n.text for n in slide.iter(A + 't')].count(expected) != 1:
                raise ValueError('Slide title missing or ambiguous')
            fields = list(slide.iter(A + 'fld'))
            if index == position:
                if len(fields) != 1 or fields[0].get('id', '').upper() != '{11111111-2222-3333-4444-555555555555}' or fields[0].get('type') != 'slidenum' or fields[0].findtext(A + 't') != str(17 + position):
                    raise ValueError('Field identity, type or cache changed')
                texts = [n.text for n in slide.iter(A + 't')]
                if texts.count('Edited page: ') != 1 or texts.count(' / note') != 1:
                    raise ValueError('Adjacent text changed')
            elif fields:
                raise ValueError('Unexpected field on a control slide')
        with zipfile.ZipFile(io.BytesIO(deck.read(target['workbookPart']))) as book:
            sheet = ET.fromstring(book.read(target['series'][0]['cells'][1]['worksheetPart']))
            cells = [c for c in sheet.iter(S + 'c') if c.get('r') == 'B3']
            if len(cells) != 1 or cells[0].findtext(S + 'v') != '7':
                raise ValueError('Target workbook value changed')
    return {'sha256': hashlib.sha256(raw).hexdigest(), 'position': position + 1, 'start': 17, 'cache': 7, 'cell': 7, 'field_cache': 17 + position, 'charts': 4}


if __name__ == '__main__':
    print(json.dumps(verify(sys.argv[1], int(sys.argv[2])), sort_keys=True))
