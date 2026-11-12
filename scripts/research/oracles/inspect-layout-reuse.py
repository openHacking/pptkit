"""Inspect already captured layout-reuse PPTX files; never operate an office app."""
import hashlib
import json
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
O = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'


def inspect(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Duplicate package parts')

        def relationships(owner):
            folder, name = posixpath.split(owner)
            part = posixpath.join(folder, '_rels', name + '.rels')
            root = ET.fromstring(archive.read(part))
            result = {}
            for item in root.findall(f'{{{R}}}Relationship'):
                identity = item.get('Id')
                if not identity or identity in result:
                    raise ValueError('Duplicate or missing relationship ID')
                target = item.get('Target')
                mode = item.get('TargetMode', 'Internal')
                result[identity] = {'type': item.get('Type', '').split('/')[-1], 'mode': mode,
                                    'target': posixpath.normpath(posixpath.join(folder, target)) if mode == 'Internal' else target}
            return result

        def single(owner, kind):
            found = [value['target'] for value in relationships(owner).values() if value['type'] == kind and value['mode'] == 'Internal']
            if len(found) != 1 or found[0] not in names:
                raise ValueError(f'Missing or ambiguous {kind} relationship on {owner}')
            return found[0]

        presentation = 'ppt/presentation.xml'
        root = ET.fromstring(archive.read(presentation))
        refs = relationships(presentation)
        pages = []
        for item in root.findall(f'{{{P}}}sldIdLst/{{{P}}}sldId'):
            reference = refs[item.get(f'{{{O}}}id')]
            if reference['type'] != 'slide' or reference['mode'] != 'Internal':
                raise ValueError('Invalid slide relationship')
            slide = reference['target']
            layout = single(slide, 'slideLayout')
            notes = single(slide, 'notesSlide')
            if single(notes, 'slide') != slide:
                raise ValueError('Incorrect notes backlink')
            node = ET.fromstring(archive.read(slide))
            links = []
            slide_rels = relationships(slide)
            for link in node.iter(f'{{{A}}}hlinkClick'):
                identity = link.get(f'{{{O}}}id')
                links.append(slide_rels.get(identity))
            pages.append({'slide': slide, 'layout': layout, 'master': single(layout, 'slideMaster'),
                          'notes': notes, 'text': ''.join(x.text or '' for x in node.iter(f'{{{A}}}t')),
                          'notesText': ''.join(x.text or '' for x in ET.fromstring(archive.read(notes)).iter(f'{{{A}}}t')),
                          'links': links})
        media = sorted(hashlib.sha256(archive.read(name)).hexdigest() for name in names if name.startswith('ppt/media/') and not name.endswith('/'))
    return {'file': str(path.resolve()), 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'pages': pages, 'mediaSha256': media}


if len(sys.argv) < 2:
    raise SystemExit('Usage: inspect-layout-reuse.py INPUT.pptx [SAVED.pptx ...]')
print(json.dumps({'oracle': 'Python zipfile and ElementTree; no native UI observation', 'files': [inspect(Path(name)) for name in sys.argv[1:]]}, indent=2))
