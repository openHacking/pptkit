#!/usr/bin/env python3
"""Inspect the retained native notes edit and its slide/notes relationship graph."""
import hashlib
import json
import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
SOURCE = ROOT / 'tests/fixtures/corpus-text-powerpoint-interactive.pptx'
OUTPUT = ROOT / 'tests/fixtures/corpus-text-notes-powerpoint-interactive.pptx'
RECORD = ROOT / 'docs/testing/corpus-notes-edit-powerpoint-evidence.json'


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def rels(z, part):
    directory, name = posixpath.split(part)
    path = posixpath.join(directory, '_rels', name + '.rels')
    return list(ET.fromstring(z.read(path)))


def target(part, rel):
    if rel.get('TargetMode') == 'External':
        raise ValueError('unexpected external relationship')
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), rel.get('Target')))


def one(z, part, kind):
    matches = [rel for rel in rels(z, part) if rel.get('Type', '').endswith('/' + kind)]
    if len(matches) != 1:
        raise ValueError(f'expected one {kind} relationship from {part}')
    return target(part, matches[0])


def inspect(path, expected):
    with ZipFile(path) as z:
        presentation = ET.fromstring(z.read('ppt/presentation.xml'))
        relationships = {rel.get('Id'): rel for rel in rels(z, 'ppt/presentation.xml')}
        slides = [target('ppt/presentation.xml', relationships[node.get(f'{{{R}}}id')]) for node in presentation.findall(f'.//{{{P}}}sldId')]
        if len(slides) != 2:
            raise ValueError('unexpected slide count')
        first = ET.fromstring(z.read(slides[0])); second = ET.fromstring(z.read(slides[1]))
        if second.get('show') != '0':
            raise ValueError('second slide is not hidden')
        slide_text = [''.join(n.text or '' for n in p.iter(f'{{{A}}}t')) for p in first.iter(f'{{{A}}}p')]
        if [s for s in slide_text if s] != ['Edited & <editable> text', 'Revised paragraph', 'Bulleted paragraph', 'Go to notes']:
            raise ValueError('first slide text changed')
        notes1, notes2 = [one(z, slide, 'notesSlide') for slide in slides]
        if notes1 == notes2 or one(z, notes1, 'slide') != slides[0] or one(z, notes2, 'slide') != slides[1]:
            raise ValueError('notes/slide ownership differs')
        first_notes = ''.join(n.text or '' for n in ET.fromstring(z.read(notes1)).iter(f'{{{A}}}t'))
        notes = ''.join(n.text or '' for n in ET.fromstring(z.read(notes2)).iter(f'{{{A}}}t'))
        if first_notes or notes != expected:
            raise ValueError('notes text differs')
        links = [node for node in first.iter(f'{{{A}}}hlinkClick') if node.get('action') == 'ppaction://hlinksldjump']
        by_id = {rel.get('Id'): rel for rel in rels(z, slides[0])}
        if len(links) != 1 or target(slides[0], by_id[links[0].get(f'{{{R}}}id')]) != slides[1]:
            raise ValueError('slide link target differs')
        return {'slides': slides, 'notes': [notes1, notes2], 'notesText': notes, 'hiddenSlide': slides[1]}


def verify(local=False):
    record = json.loads(RECORD.read_text())
    for path, key in ((SOURCE, 'source'), (OUTPUT, 'output')):
        if digest(path) != record[key]['sha256']:
            raise ValueError(f'{key} hash mismatch')
    inspect(SOURCE, 'Speaker & reviewer notes')
    inspect(OUTPUT, 'Speaker & reviewer notes native edit')
    if local:
        directory = ROOT / record['artifactDirectory']
        if digest(directory / 'text-notes-edit.pptx') != record['output']['sha256']:
            raise ValueError('local PowerPoint output hash mismatch')
        for name, sha in record['captures'].items():
            if digest(directory / name) != sha:
                raise ValueError(f'capture hash mismatch: {name}')
        for name in ('notes-edit-first.ax.txt', 'notes-edit-reopen.ax.txt'):
            text = (directory / name).read_text()
            if 'Speaker & reviewer notes native edit' not in text or '幻灯片 2 / 2' not in text or '隐藏' not in text:
                raise ValueError(f'UI state mismatch: {name}')


if __name__ == '__main__':
    import sys
    if sys.argv[1:] not in ([], ['--local']):
        raise SystemExit('usage: verify-corpus-notes-edit.py [--local]')
    verify(bool(sys.argv[1:]))
    print('verified notes edit package' + (' and local captures' if sys.argv[1:] else ''))
