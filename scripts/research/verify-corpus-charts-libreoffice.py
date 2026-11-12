#!/usr/bin/env python3
"""Verify the saved native chart edit, independently of the LibreOffice UI."""
import hashlib
import importlib.util
import json
import posixpath
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
NS = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}


def check_saved(path):
    expected = {'bar': ('barChart', ('1.5', '4', '3')), 'line': ('lineChart', ('2', '4', '3')),
                'pie': ('pieChart', ('2', '4', '3'))}
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(set(names)) != len(names) or archive.testzip() is not None:
            raise ValueError('Invalid or duplicate ZIP entry')
        if any(name.startswith('ppt/embeddings/') and name.endswith('.xlsx') for name in names):
            raise ValueError('Unexpected embedded worksheet')
        slide = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
        rels = ET.fromstring(archive.read('ppt/slides/_rels/slide1.xml.rels'))
        seen = set()
        for binding in slide.findall('.//c:chart', NS):
            rid = binding.attrib[f'{{{NS["r"]}}}id']
            matches = [r for r in rels.findall('rel:Relationship', NS) if r.get('Id') == rid]
            if len(matches) != 1 or not matches[0].get('Type', '').endswith('/chart') or matches[0].get('TargetMode') == 'External':
                raise ValueError('Unexpected chart relationship')
            part = posixpath.normpath(posixpath.join('ppt/slides', matches[0].attrib['Target']))
            chart = ET.fromstring(archive.read(part))
            title = ''.join(t.text or '' for t in chart.findall('.//c:title//a:t', NS))
            if title not in expected or title in seen or chart.find('c:externalData', NS) is not None:
                raise ValueError('Unexpected chart identity or external data declaration')
            seen.add(title)
            kind, values = expected[title]
            series = chart.findall(f'c:chart/c:plotArea/c:{kind}/c:ser', NS)
            if len(series) != 1:
                raise ValueError(f'Unexpected series in {title}')
            points = series[0].findall('c:val/c:numRef/c:numCache/c:pt', NS)
            if [(p.get('idx'), p.findtext('c:v', namespaces=NS)) for p in points] != list(zip(('0', '1', '2'), values)):
                raise ValueError(f'Unexpected chart values in {title}')
        if seen != set(expected):
            raise ValueError('Slide does not reference all three expected charts')


def verify(local=False):
    evidence = json.loads((ROOT / 'docs/testing/corpus-charts-libreoffice-interactive-evidence.json').read_text())
    output = ROOT / evidence['output']['path']
    if hashlib.sha256(output.read_bytes()).hexdigest() != evidence['output']['sha256']:
        raise ValueError('Saved output hash differs')
    check_saved(output)
    if local:
        source = ROOT / evidence['input']['path']
        if hashlib.sha256(source.read_bytes()).hexdigest() != evidence['input']['sha256']:
            raise ValueError('Source hash differs')
        spec = importlib.util.spec_from_file_location('source_check', Path(__file__).with_name('verify-corpus-charts-powerpoint.py'))
        source_check = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(source_check)
        source_check.inspect(source, False)


if __name__ == '__main__':
    if sys.argv[1:] not in ([], ['--local']):
        raise SystemExit('usage: verify-corpus-charts-libreoffice.py [--local]')
    verify(bool(sys.argv[1:]))
    print('Verified saved native chart values and control charts' + ('; verified original workbooks' if sys.argv[1:] else ''))
