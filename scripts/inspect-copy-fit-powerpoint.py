#!/usr/bin/env python3
"""Inspect previously captured PowerPoint fit files; does not operate the application."""
import hashlib
import json
import platform
import posixpath
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

BASE = Path('.tmp/copy-fit-powerpoint')
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'

def sha(file):
    return hashlib.sha256(file.read_bytes()).hexdigest()

def inspect(file):
    with zipfile.ZipFile(file) as z:
        slide = ET.fromstring(z.read('ppt/slides/slide1.xml'))
        notes = ET.fromstring(z.read('ppt/notesSlides/notesSlide1.xml'))
        def text(root):
            return ' '.join(n.text or '' for n in root.iter(f'{{{A}}}t'))
        def relations(part):
            folder, name = posixpath.split(part)
            root = ET.fromstring(z.read(f'{folder}/_rels/{name}.rels'))
            return [{'type': x.get('Type', '').split('/')[-1], 'mode': x.get('TargetMode', 'Internal'),
                     'target': posixpath.normpath(posixpath.join(folder, x.get('Target')))} for x in root if x.tag == f'{{{R}}}Relationship']
        slide_rel = relations('ppt/slides/slide1.xml'); notes_rel = relations('ppt/notesSlides/notesSlide1.xml')
        assert any(r['type'] == 'notesSlide' and r['mode'] == 'Internal' and r['target'] == 'ppt/notesSlides/notesSlide1.xml' for r in slide_rel)
        assert any(r['type'] == 'slide' and r['mode'] == 'Internal' and r['target'] == 'ppt/slides/slide1.xml' for r in notes_rel)
        def node(element):
            kind = element.tag.split('}')[-1]
            props = element.find(f'.//{{{P}}}cNvPr')
            trans = element.find(f'{{{P}}}spPr/{{{A}}}xfrm') if kind in ('sp', 'pic', 'cxnSp') else element.find(f'{{{P}}}grpSpPr/{{{A}}}xfrm') if kind == 'grpSp' else element.find(f'{{{P}}}xfrm')
            geom = {}
            if trans is not None:
                for key in ('off', 'ext', 'chOff', 'chExt'):
                    value = trans.find(f'{{{A}}}{key}')
                    if value is not None: geom[key] = dict(value.attrib)
            return {'kind': kind, 'name': props.get('name') if props is not None else None, 'geometry': geom,
                    'text': text(element), 'grid': [int(x.get('w')) for x in element.iter(f'{{{A}}}gridCol')],
                    'rows': [int(x.get('h')) for x in element.iter(f'{{{A}}}tr')],
                    'children': [node(x) for x in element if x.tag.split('}')[-1] in ('sp', 'pic', 'cxnSp', 'grpSp', 'graphicFrame')] if kind == 'grpSp' else []}
        tree = slide.find(f'.//{{{P}}}spTree')
        objects = [node(x) for x in tree if x.tag.split('}')[-1] in ('sp', 'pic', 'cxnSp', 'grpSp', 'graphicFrame')]
        assert all(word in text(slide) for word in ('Geometry fit', 'Left', 'Right')) and 'Fit notes' in text(notes)
        assert len(objects) == 5
        return {'file': str(file.resolve()), 'sha256': sha(file), 'slideText': text(slide), 'notesText': text(notes), 'objects': objects, 'slideRelationships': slide_rel, 'notesRelationships': notes_rel}

files = [BASE / 'fitted-input.pptx'] + [BASE / f'fit-round{i}.pptx' for i in (1, 2, 3)]
rounds = [inspect(f) for f in files]
assert rounds[0]['objects'][1]['geometry']['off'] == {'x': '127000', 'y': '889000'}
assert rounds[0]['objects'][3]['geometry']['ext'] == {'cx': '635000', 'cy': '254000'}
assert rounds[0]['objects'][4]['grid'] == [381000, 635000] and rounds[0]['objects'][4]['rows'] == [254000]
for item in rounds[1:]:
    assert [o['geometry'] for o in item['objects'][:4]] == [o['geometry'] for o in rounds[0]['objects'][:4]]
    assert item['objects'][3]['children'][0]['geometry'] == rounds[0]['objects'][3]['children'][0]['geometry']
    assert item['objects'][4]['grid'] == [381000, 635000] and item['objects'][4]['rows'] == [254000]
    assert item['objects'][4]['geometry']['ext'] == {'cx': '1016000', 'cy': '911860'}
shots = [BASE / f for f in ('control-first.png', 'fit-first.png', 'fit-round1.png', 'fit-round2.png', 'fit-round3.png', 'fit-round3-notes.png')]
version = subprocess.check_output(['defaults', 'read', '/Applications/Microsoft PowerPoint.app/Contents/Info', 'CFBundleShortVersionString'], text=True).strip()
record = {'schemaVersion': 1, 'scope': 'One source and one explicit-geometry fitted page; PowerPoint for Mac three successive Save As and reopen observations. Static connector, no link activation or interactive edit claim.',
          'environment': {'application': 'Microsoft PowerPoint ' + version, 'system': platform.platform(), 'architecture': platform.machine()},
          'observation': 'Source page text was continuous; fitted input and each reopened file visibly wrapped text and overlapped the red shape. Third reopen displayed Fit notes. No repair prompt was observed.',
          'sourceControl': {'file': str((BASE / 'source-control.pptx').resolve()), 'sha256': sha(BASE / 'source-control.pptx')},
          'independent': {'oracle': 'Python zipfile and ElementTree', 'rounds': rounds},
          'screenshots': [{'file': str(f.resolve()), 'sha256': sha(f)} for f in shots]}
Path('docs/testing/copy-fit-powerpoint-evidence.json').write_text(json.dumps(record, indent=2) + '\n')
print(json.dumps({'application': version, 'rounds': len(rounds), 'tableFrameHeights': [r['objects'][4]['geometry']['ext']['cy'] for r in rounds]}, indent=2))
