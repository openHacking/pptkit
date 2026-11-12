#!/usr/bin/env python3
"""Independently re-read retained Calc workbook cells and reject forged evidence."""
import hashlib
import json
import re
import sys
import xml.etree.ElementTree as E
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
EXPECTED = {'B1': 'name _x000D_', 'A2': '_x0041_', 'A3': 'line\rbreak', 'A4': '_x005F_x0041_', 'A5': 'plain & 中文', 'A6': '_x0041__x0042_'}
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def decode(text): return re.sub(r'_x([0-9a-fA-F]{4})_', lambda m: chr(int(m.group(1), 16)), text)

def verify(manifest):
    evidence = json.loads(manifest.read_text())
    assert evidence['runnerSha256'] == sha(ROOT / 'scripts/research/chart-xstring-calc.py')
    assert evidence['generatorSha256'] == sha(ROOT / 'scripts/prepare-chart-xstring-workbook.mjs')
    assert len(evidence['records']) == 4 and len(evidence['logs']) == 3
    observations = []
    for index, record in enumerate(evidence['records']):
        assert record['round'] == index
        path = (ROOT / record['path']).resolve(); assert path.is_relative_to((ROOT / '.tmp/chart-xstring-calc').resolve())
        assert sha(path) == record['sha256']
        with zipfile.ZipFile(path) as book:
            assert len(book.namelist()) == len(set(book.namelist()))
            strings = E.fromstring(book.read('xl/sharedStrings.xml')) if 'xl/sharedStrings.xml' in book.namelist() else None
            shared = [''.join(x.itertext()) for x in strings.findall(S + 'si')] if strings is not None else []
            sheet = E.fromstring(book.read('xl/worksheets/sheet1.xml'))
            cells = {}
            for cell in sheet.iter(S + 'c'):
                ref = cell.get('r'); assert ref not in cells
                kind = cell.get('t', 'n')
                raw = shared[int(cell.findtext(S + 'v'))] if kind == 's' else ''.join(cell.find(S + 'is').itertext()) if kind == 'inlineStr' else cell.findtext(S + 'v')
                cells[ref] = {'type': kind, 'raw': raw}
            result = {}
            for ref in EXPECTED:
                value = cells[ref]; value['decodedOnce'] = decode(value['raw']); assert value == record['cells'][ref]; result[ref] = value['decodedOnce']
            observations.append(result)
    assert observations[0] == EXPECTED
    for result in observations[1:]:
        for ref, expected in EXPECTED.items(): assert result[ref] == ('_x0041_' if ref == 'A4' else expected)
    assert all(x['exitCode'] == 0 for x in evidence['logs'])
    return {'verified': True, 'records': len(observations), 'changedCell': 'A4', 'source': EXPECTED['A4'], 'resaved': observations[-1]['A4']}

if __name__ == '__main__': print(json.dumps(verify(Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/chart-xstring-calc-evidence.json'), ensure_ascii=False))
