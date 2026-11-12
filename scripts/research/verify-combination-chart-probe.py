#!/usr/bin/env python3
"""Verify axis graph, cached data, workbook cells, and native plotted geometry."""
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
CATEGORIES = ['Jan', 'Feb', 'Mar']
VALUES = [[2, 5, 8], [800, 200, 600]]


def check(condition, message):
    if not condition:
        raise ValueError(message)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def package(path, source, values=VALUES, *, powerpoint=False):
    with ZipFile(path) as archive:
        check(archive.testzip() is None, 'PPTX CRC failed')
        chart = ET.fromstring(archive.read('ppt/charts/chart1.xml'))
        plot = chart.find(f'.//{C}plotArea')
        scalar = lambda node, tag: node.find(f'{C}{tag}').get('val')
        axes = {}
        for node in plot:
            kind = node.tag.removeprefix(C)
            if kind not in ('catAx', 'valAx'):
                continue
            ident = scalar(node, 'axId')
            check(ident not in axes, 'duplicate axis id')
            axes[ident] = {'kind': kind, 'position': scalar(node, 'axPos'), 'cross': scalar(node, 'crossAx'),
                           'deleted': scalar(node, 'delete') in ('1', 'true')}
            if kind == 'valAx':
                scaling = node.find(f'{C}scaling')
                axes[ident]['bounds'] = [int(scalar(scaling, tag)) for tag in ('min', 'max')]
        check(len(axes) == 4, 'expected two distinct axis pairs')
        groups = []
        for i, kind in enumerate(('barChart', 'lineChart')):
            group, = plot.findall(f'{C}{kind}')
            refs = [node.get('val') for node in group.findall(f'{C}axId')]
            check(len(refs) == 2 and all(ref in axes for ref in refs), 'chart family axis references differ')
            category, value = [axes[ref] for ref in refs]
            check(category['kind'] == 'catAx' and value['kind'] == 'valAx', 'wrong axis types')
            check(category['cross'] == refs[1] and value['cross'] == refs[0], 'cross axes are not reciprocal')
            check(category['position'] == ('b', 'b' if powerpoint else 't')[i] and value['position'] == ('l', 'r')[i], 'axis placement differs')
            check(category['deleted'] == bool(i) and not value['deleted'], 'axis visibility differs')
            check(value['bounds'] == [0, (10, 1000)[i]], 'value axis bounds differ')
            series, = group.findall(f'{C}ser')
            column = ('B', 'C')[i]
            cats, nums = [series.find(f'{C}{key}/{C}{ref}') for key, ref in [('cat', 'strRef'), ('val', 'numRef')]]
            cache = lambda node, kind: [(int(point.attrib['idx']), point.findtext(f'{C}v')) for point in node.findall(f'{C}{kind}/{C}pt')]
            check(cache(cats, 'strCache') == list(enumerate(CATEGORIES)), 'category order differs')
            check(cache(nums, 'numCache') == list(enumerate(map(str, values[i]))), 'numeric cache differs')
            if source:
                check(cats.findtext(f'{C}f') == 'Sheet1!$A$2:$A$4' and nums.findtext(f'{C}f') == f'Sheet1!${column}$2:${column}$4', 'source workbook ranges differ')
            groups.append({'kind': kind, 'axes': refs, 'values': values[i]})
        books = [name for name in archive.namelist() if name.endswith('.xlsx')]
        external = chart.find(f'{C}externalData')
        if source:
            check(len(books) == 1 and external is not None, 'source workbook missing')
            rels = ET.fromstring(archive.read('ppt/charts/_rels/chart1.xml.rels'))
            rel, = [r for r in rels if r.get('Id') == external.get(f'{R}id')]
            check(posixpath.normpath(posixpath.join('ppt/charts', rel.get('Target'))) == books[0], 'wrong workbook target')
            with ZipFile(io.BytesIO(archive.read(books[0]))) as book:
                sheet = ET.fromstring(book.read('xl/worksheets/sheet1.xml'))
                cells = {cell.attrib['r']: cell for cell in sheet.iter(f'{S}c')}
                for i, category in enumerate(CATEGORIES, 2):
                    check(cells[f'A{i}'].findtext(f'{S}is/{S}t') == category, 'wrong category cell')
                    check([int(cells[f'{col}{i}'].findtext(f'{S}v')) for col in ('B', 'C')] == [v[i - 2] for v in values], 'wrong numeric cells')
        else:
            check(not books and external is None, 'native workbook outcome changed')
    return {'sha256': sha(path), 'axes': axes, 'series': groups, 'embeddedWorkbooks': len(books)}


def pdf(path, values=VALUES):
    def rgb(color):
        return tuple(round(x * 255) for x in color) if color else None
    with fitz.open(path) as document:
        check(len(document) == 1, 'expected one visible page')
        page = document[0]
        words = page.get_text('words')
        left = {word[4]: (word[1] + word[3]) / 2 for word in words if word[0] < 120 and word[4].isdigit()}
        right = {word[4]: (word[1] + word[3]) / 2 for word in words if word[0] > page.rect.width * .75 and word[4].isdigit()}
        check(set(left) == {'0', '2', '4', '6', '8', '10'}, 'left tick labels differ')
        check(set(right) == {'0', '200', '400', '600', '800', '1000'}, 'right tick labels differ')
        check(all(abs(left[str(value)] - right[str(value * 100)]) < .5 for value in range(0, 11, 2)), 'axis label scales are not aligned')
        bars = sorted([shape['rect'] for shape in page.get_drawings() if rgb(shape['fill']) == (68, 114, 196)], key=lambda rect: rect.x0)
        line, = [shape for shape in page.get_drawings() if rgb(shape['color']) == (237, 125, 49) and len(shape['items']) == 2 and all(item[0] == 'l' for item in shape['items'])]
        points = [line['items'][0][1], *(item[2] for item in line['items'])]
        check(len(bars) == len(points) == 3, 'plot paths differ')
        height = left['0'] - left['10']
        zero = bars[0].y1
        check(all(abs(bar.y1 - zero) < .5 and abs((bar.x0 + bar.x1) / 2 - point.x) < .5 for bar, point in zip(bars, points)), 'bar and line category centres differ')
        normalized = [[(zero - bar.y0) / height for bar in bars], [(zero - point.y) / height for point in points]]
        for actual, series, maximum in zip(normalized, values, (10, 1000)):
            expected = [value / maximum for value in series]
            check(all(abs(a - b) < .005 for a, b in zip(actual, expected)), 'plot values do not follow the separate axis scales')
    return {'sha256': sha(path), 'barRectsPt': [list(rect) for rect in bars], 'linePointsPt': [list(point) for point in points],
            'leftTicks': left, 'rightTicks': right, 'normalizedBarAndLineHeights': normalized}


if __name__ == '__main__':
    root = Path(sys.argv[1])
    if '--powerpoint-resave' in sys.argv[2:]:
        edited_values = [[2, 7, 8], [800, 400, 600]]
        print(json.dumps(package(root, True, edited_values, powerpoint=True), indent=2))
        sys.exit(0)
    scope = 'Public API native combination probe' if '--public' in sys.argv[2:] else 'Internal native combination probe'
    edited = '--edited' in sys.argv[2:]
    values = [[2, 7, 8], [800, 400, 600]] if edited else VALUES
    filename = 'combination-edited' if edited else 'combination'
    if edited:
        scope = 'Combination source edit-save probe (Feb: columns 5 to 7, lines 200 to 400)'
    report = {'schemaVersion': 1, 'scope': f'{scope}: headless open, save and reopen only.',
              'directory': str(root), 'application': (root / 'version.txt').read_text().strip(),
              'platform': platform.platform(), 'pymupdf': fitz.VersionBind, 'source': package(root / f'{filename}.pptx', True, values),
              'saved': package(root / f'saved/{filename}.pptx', False, values), 'firstPdf': pdf(root / f'pdf/{filename}.pdf', values),
              'reopenedPdf': pdf(root / f'reopened/{filename}.pdf', values),
              'logsSha256': {name: sha(root / name) for name in ('open.log', 'save.log', 'reopen.log')}}
    print(json.dumps(report, indent=2))
