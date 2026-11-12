#!/usr/bin/env python3
"""Verify the CLI bubble source edit and LibreOffice's resulting PDF geometry."""
import importlib.util
import io
import json
import platform
import re
import sys
from pathlib import Path
from zipfile import ZipFile

spec = importlib.util.spec_from_file_location('bubble_probe', Path(__file__).with_name('verify-bubble-export-probe.py'))
probe = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe)
EXPECTED = [[['1', '5'], ['2', '5'], ['1', '16']], [['7'], ['6'], ['4']]]
CHART = 'ppt/charts/chart1.xml'
BOOK = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'
SHEET = 'xl/worksheets/sheet1.xml'


def verify_edit(source, edited):
    with ZipFile(source) as before, ZipFile(edited) as after:
        names = before.namelist()
        probe.require(sorted(names) == sorted(after.namelist()), 'PPTX entries changed')
        for name in names:
            if name not in (CHART, BOOK):
                probe.require(before.read(name) == after.read(name), f'unrelated PPTX payload changed: {name}')
        chart = before.read(CHART)
        for tag, old, new in [(b'xVal', b'4', b'5'), (b'yVal', b'4', b'5'), (b'bubbleSize', b'9', b'16')]:
            pattern = rb'(<c:' + tag + rb'>[^<]*(?:(?!</c:' + tag + rb'>).)*?<c:pt idx="1"><c:v>)' + old + rb'(</c:v>)'
            chart, count = re.subn(pattern, lambda m: m[1] + new + m[2], chart, count=1, flags=re.S)
            probe.require(count == 1, f'unexpected original cache: {tag}')
        probe.require(chart == after.read(CHART), 'chart changed outside the three selected cache points')
        with ZipFile(io.BytesIO(before.read(BOOK))) as old_book, ZipFile(io.BytesIO(after.read(BOOK))) as new_book:
            names = old_book.namelist()
            probe.require(sorted(names) == sorted(new_book.namelist()), 'worksheet entries changed')
            for name in names:
                old = old_book.read(name)
                if name == SHEET:
                    for cell, prior, new in [(b'A3', b'4', b'5'), (b'B3', b'4', b'5'), (b'C3', b'9', b'16')]:
                        prefix = b'<c r="' + cell + b'"><v>'
                        original = prefix + prior + b'</v></c>'
                        probe.require(old.count(original) == 1, f'unexpected original worksheet cell: {cell}')
                        old = old.replace(original, prefix + new + b'</v></c>', 1)
                probe.require(old == new_book.read(name), f'unexpected XLSX payload change: {name}')


def geometry(path, x, y, ratios):
    result = probe.pdf(path, ratios)
    centers = [((a + c) / 2, (b + d) / 2) for a, b, c, d in [bubble['rectPt'] for bubble in result['bubbles']]]
    low, selected, high = centers
    # Both unchanged bubbles are reference points on the two fixed linear axes.
    # Their authored centres are (1, 2) and (7, 6).
    observed = [1 + 6 * (selected[0] - low[0]) / (high[0] - low[0]),
                2 + 4 * (selected[1] - low[1]) / (high[1] - low[1])]
    probe.require(all(abs(a - b) < .01 for a, b in zip(observed, (x, y))), 'selected bubble center differs')
    return {**result, 'selectedCenterInDataUnits': observed}


def verify(root):
    source, edited = root / 'source.pptx', root / 'bubble.pptx'
    evidence = json.loads((Path(__file__).resolve().parents[2] / 'docs/testing/bubble-public-probe-evidence.json').read_text())
    original = probe.package(source, True)
    probe.require(original['sha256'] == evidence['source']['sha256'], 'source differs from public probe')
    verify_edit(source, edited)
    return {
        'schemaVersion': 1, 'scope': 'PPTKit CLI source edit followed by LibreOffice headless open, save and reopen; no native data editing.',
        'directory': str(root), 'application': (root / 'version.txt').read_text().strip(), 'platform': platform.platform(),
        'pymupdf': probe.fitz.VersionBind, 'source': original,
        'cliEdited': probe.package(edited, True, EXPECTED),
        'nativeSaved': probe.package(root / 'saved/bubble.pptx', False, EXPECTED),
        'sourcePdf': geometry(root / 'pdf/source.pdf', 4, 4, (3, 2)),
        'cliEditedPdf': geometry(root / 'pdf/bubble.pdf', 5, 5, (4, 2)),
        'nativeReopenedPdf': geometry(root / 'reopened/bubble.pdf', 5, 5, (4, 2)),
        'changedAtoms': {'chart': ['X[1]', 'Y[1]', 'size[1]'], 'worksheet': ['A3', 'B3', 'C3']},
        'logsSha256': {name: probe.digest(root / name) for name in ('edits.json', 'cli-save.json', 'native-open.log', 'native-save.log', 'native-reopen.log')},
    }


if __name__ == '__main__':
    print(json.dumps(verify(Path(sys.argv[1])), indent=2))
