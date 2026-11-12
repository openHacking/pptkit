#!/usr/bin/env python3
"""Observe fresh chart workbook strings through three isolated headless Calc resaves."""
import hashlib
import json
import platform
import re
import subprocess
import tempfile
import xml.etree.ElementTree as E
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
BASE = ROOT / '.tmp/chart-xstring-calc'

def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def decode(value): return re.sub(r'_x([0-9a-fA-F]{4})_', lambda m: chr(int(m.group(1), 16)), value)

def inventory(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)): raise ValueError('duplicate ZIP entries')
        shared = []
        if 'xl/sharedStrings.xml' in names:
            root = E.fromstring(archive.read('xl/sharedStrings.xml'))
            shared = [''.join(node.itertext()) for node in root.findall(S + 'si')]
        root = E.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        cells = {}
        for cell in root.iter(S + 'c'):
            ref, kind = cell.get('r'), cell.get('t', 'n')
            if ref in cells: raise ValueError('duplicate worksheet cell')
            if kind == 's':
                value = shared[int(cell.findtext(S + 'v'))]
            elif kind == 'inlineStr':
                value = ''.join(cell.find(S + 'is').itertext())
            else: value = cell.findtext(S + 'v')
            cells[ref] = {'type': kind, 'raw': value}
        return cells

def main():
    BASE.mkdir(parents=True, exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix='run-', dir=BASE)); source = run / 'source.xlsx'
    subprocess.run(['node', 'scripts/prepare-chart-xstring-workbook.mjs', str(source)], cwd=ROOT, check=True)
    records, logs = [], []
    expected = {'B1': 'name _x005F_x000D_', 'A2': '_x005F_x0041_', 'A3': 'line_x000D_break',
                'A4': '_x005F_x005F_x005F_x0041_', 'A5': 'plain & 中文', 'A6': '_x005F_x0041__x005F_x0042_'}
    for round in range(4):
        path = source if round == 0 else run / f'round-{round}' / 'source.xlsx'
        if round:
            path.parent.mkdir()
            previous = source if round == 1 else run / f'round-{round-1}' / 'source.xlsx'
            process = subprocess.run([str(SOFFICE), f'-env:UserInstallation={(run / f"profile-{round}").as_uri()}', '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(path.parent), str(previous)], capture_output=True, text=True, timeout=60)
            logs.append({'round': round, 'exitCode': process.returncode, 'stdout': process.stdout, 'stderr': process.stderr})
            if process.returncode or not path.is_file(): raise ValueError(f'Calc conversion failed: {process.stderr}')
        cells = inventory(path)
        if round == 0 and {k: cells[k]['raw'] for k in expected} != expected: raise ValueError('source lexical sample disagrees')
        records.append({'round': round, 'path': str(path.relative_to(ROOT)), 'sha256': digest(path), 'cells': {k: {**cells[k], 'decodedOnce': decode(cells[k]['raw'])} for k in expected}})
    evidence = {'scope': 'One authored chart workbook, three LibreOfficeDev Calc headless XLSX resaves; raw XML values only, no PowerPoint, GUI or chart cache.',
                'application': subprocess.run([str(SOFFICE), '--version'], capture_output=True, text=True, check=True).stdout.strip(),
                'platform': platform.platform(), 'runnerSha256': digest(Path(__file__)), 'generatorSha256': digest(ROOT / 'scripts/prepare-chart-xstring-workbook.mjs'), 'records': records, 'logs': logs}
    (ROOT / 'docs/testing/chart-xstring-calc-evidence.json').write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps(evidence, ensure_ascii=False))

if __name__ == '__main__': main()
