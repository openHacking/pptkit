#!/usr/bin/env python3
"""Read captured PowerPoint files; never starts or controls PowerPoint."""
import hashlib
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

directory = Path('.tmp/native-field-powerpoint-01')
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
records = []
for name in ('source.pptx', 'round1.pptx', 'round2.pptx', 'round3.pptx'):
    path = directory / name
    raw = path.read_bytes()
    slides = []
    with zipfile.ZipFile(path) as archive:
        for index in (1, 2):
            root = ET.fromstring(archive.read(f'ppt/slides/slide{index}.xml'))
            fields = [{'id': field.get('id'), 'type': field.get('type'),
                       'cache': ''.join(field.itertext())}
                      for field in root.findall('.//a:fld', ns)]
            expected = [
                {'id': f'{{11111111-2222-3333-4444-{index:012d}}}', 'type': 'slidenum', 'cache': 'STALE-NUMBER' if name == 'source.pptx' else str(index)},
                {'id': f'{{AAAAAAAA-BBBB-CCCC-DDDD-{index:012d}}}', 'type': 'datetime3', 'cache': 'STALE-DATE' if name == 'source.pptx' else '28 October 2026'},
            ]
            if fields != expected:
                raise ValueError(f'{name} slide {index}: {fields!r} != {expected!r}')
            slides.append(fields)
    records.append({'file': str(path), 'sha256': hashlib.sha256(raw).hexdigest(), 'slides': slides})
evidence = json.loads(Path('docs/testing/native-fields-powerpoint-evidence.json').read_text())
if records != evidence['records']:
    raise ValueError('Captured file hashes or field records differ from the evidence manifest')
print(json.dumps(records, ensure_ascii=False, indent=2))
