#!/usr/bin/env python3
"""Check captured files in presentation order without launching PowerPoint."""
import hashlib
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

base = Path('.tmp/native-field-powerpoint-01')
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
records = []
for name in ('reordered.pptx', 'edited.pptx'):
    path = base / name
    raw = path.read_bytes()
    slides = []
    with zipfile.ZipFile(path) as archive:
        presentation = ET.fromstring(archive.read('ppt/presentation.xml'))
        relationships = ET.fromstring(archive.read('ppt/_rels/presentation.xml.rels'))
        targets = {rel.get('Id'): rel.get('Target') for rel in relationships}
        for slide in presentation.findall('./p:sldIdLst/p:sldId', ns):
            target = targets[slide.get(f'{{{ns["r"]}}}id')].lstrip('/')
            target = target if target.startswith('ppt/') else 'ppt/' + target
            root = ET.fromstring(archive.read(target))
            fields = [{'id': f.get('id'), 'type': f.get('type'), 'cache': ''.join(f.itertext())}
                      for f in root.findall('.//a:fld', ns)]
            text = ''.join(node.text or '' for node in root.findall('.//a:t', ns))
            slides.append({'part': target, 'text': text, 'fields': fields})
    if len(slides) != 2:
        raise ValueError(f'{name}: expected two slides')
    for pos, original in enumerate((2, 1), 1):
        expected = [{'id': f'{{11111111-2222-3333-4444-{original:012d}}}', 'type': 'slidenum', 'cache': str(pos)},
                    {'id': f'{{AAAAAAAA-BBBB-CCCC-DDDD-{original:012d}}}', 'type': 'datetime3', 'cache': '28 October 2026'}]
        text = ('EDIT ' if name == 'edited.pptx' and pos == 1 else '') + f'Page {original}: {pos} / Date: 28 October 2026'
        if slides[pos - 1]['fields'] != expected or slides[pos - 1]['text'] != text:
            raise ValueError(f'{name} position {pos}: {slides[pos - 1]!r}')
    records.append({'file': str(path), 'sha256': hashlib.sha256(raw).hexdigest(), 'slides': slides})
manifest = Path('docs/testing/native-field-edits-powerpoint-evidence.json')
if manifest.exists() and records != json.loads(manifest.read_text())['records']:
    raise ValueError('Captured files differ from their recorded hashes or contents')
print(json.dumps(records, ensure_ascii=False, indent=2))
