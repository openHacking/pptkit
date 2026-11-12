#!/usr/bin/env python3
"""Inspect previously captured native PowerPoint files, without automating the application."""
import hashlib
import json
import platform
import posixpath
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

BASE = Path('.tmp/copy-placeholder-index-powerpoint')
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'

def sha(file):
    return hashlib.sha256(file.read_bytes()).hexdigest()

def inspect(file):
    with zipfile.ZipFile(file) as z:
        def one(part, kind):
            folder, name = posixpath.split(part)
            root = ET.fromstring(z.read(f'{folder}/_rels/{name}.rels'))
            matches = [posixpath.normpath(posixpath.join(folder, r.get('Target'))) for r in root if r.tag == f'{{{R}}}Relationship' and r.get('Type', '').endswith('/' + kind) and r.get('TargetMode', 'Internal') == 'Internal']
            return matches[0] if len(matches) == 1 else None
        def placeholders(part):
            result = []
            for sp in ET.fromstring(z.read(part)).iter(f'{{{P}}}sp'):
                ph = sp.find(f'{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph')
                name = sp.find(f'{{{P}}}nvSpPr/{{{P}}}cNvPr')
                if ph is not None:
                    result.append({'name': name.get('name') if name is not None else None, 'type': ph.get('type'), 'index': ph.get('idx'), 'text': ' '.join(t.text or '' for t in sp.iter(f'{{{A}}}t'))})
            return result
        pages = []
        for i in (1, 2):
            slide = f'ppt/slides/slide{i}.xml'; layout = one(slide, 'slideLayout'); notes = one(slide, 'notesSlide')
            assert layout and notes and one(notes, 'slide') == slide
            note = ' '.join(t.text or '' for t in ET.fromstring(z.read(notes)).iter(f'{{{A}}}t'))
            assert note == f'Notes {chr(64+i)}'
            page = {'part': slide, 'layout': layout, 'slidePlaceholders': placeholders(slide), 'layoutPlaceholders': placeholders(layout), 'notes': notes, 'notesText': note, 'notesBacklink': one(notes, 'slide')}
            assert [p['text'] for p in page['slidePlaceholders']] == [f'Title {chr(64+i)}', f'Body {chr(64+i)}']
            pages.append(page)
        assert pages[0]['layout'] != pages[1]['layout']
        return {'file': str(file.resolve()), 'sha256': sha(file), 'pages': pages}

files = [BASE / 'input.pptx'] + [BASE / f'index-round{i}.pptx' for i in (1, 2, 3)]
rounds = [inspect(f) for f in files]
for round_number, item in enumerate(rounds):
    for page in item['pages']:
        expected = [('title', '0' if round_number == 0 else None), ('body', '4294967295')]
        assert [(p['type'], p['index']) for p in page['slidePlaceholders']] == expected
        assert [(p['type'], p['index']) for p in page['layoutPlaceholders']] == expected
shots = [BASE / f'first-slide{i}-notes.png' for i in (1, 2)] + [BASE / f'round{r}-slide{i}.png' for r in (1, 2, 3) for i in (1, 2)]
version = subprocess.check_output(['defaults', 'read', '/Applications/Microsoft PowerPoint.app/Contents/Info', 'CFBundleShortVersionString'], text=True).strip()
evidence = {'schemaVersion': 1, 'scope': 'Two copied slides with per-page layout mappings and explicit indexes 0 and 4294967295; PowerPoint for Mac three successive Save As and reopen observations. No interactive placeholder edit or general compatibility claim.',
            'environment': {'application': 'Microsoft PowerPoint ' + version, 'system': platform.platform(), 'architecture': platform.machine()},
            'observation': 'Both pages and matching notes were visible at first open and every reopen; no repair prompt was observed.',
            'independent': {'oracle': 'Python zipfile and ElementTree', 'rounds': rounds},
            'screenshots': [{'file': str(f.resolve()), 'sha256': sha(f)} for f in shots]}
Path('docs/testing/copy-placeholder-index-powerpoint-evidence.json').write_text(json.dumps(evidence, indent=2) + '\n')
print(json.dumps({'application': version, 'rounds': len(rounds), 'indexes': [[[(p['type'], p['index']) for p in page['slidePlaceholders']] for page in item['pages']] for item in rounds]}, indent=2))
