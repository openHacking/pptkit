#!/usr/bin/env python3
"""Independent ZIP/XML inventory and deterministic sentinel fixture setup."""
import json
import posixpath
import sys
import zipfile
from xml.etree import ElementTree as ET

CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def core_part(archive):
    rels = ET.fromstring(archive.read('_rels/.rels'))
    found = [r for r in rels if r.get('Type') == CORE_REL]
    if len(found) != 1 or found[0].get('TargetMode') == 'External':
        raise ValueError('Core properties relationship is missing, ambiguous or external')
    target = found[0].get('Target')
    if not target or '\\' in target:
        raise ValueError('Invalid core properties target')
    part = posixpath.normpath(target.lstrip('/'))
    if part.startswith('../') or part not in archive.namelist():
        raise ValueError('Core properties target is unavailable')
    return part


def inventory(path):
    with zipfile.ZipFile(path) as archive:
        if len(archive.namelist()) != len(set(archive.namelist())):
            raise ValueError('Duplicate ZIP entries')
        part = core_part(archive)
        root = ET.fromstring(archive.read(part))
        if root.tag != '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}coreProperties':
            raise ValueError('Unexpected core properties root')
        properties = []
        for child in root:
            properties.append({'expandedName': child.tag, 'value': ''.join(child.itertext()), 'attributes': dict(sorted(child.attrib.items()))})
        slides = []
        for name in sorted(n for n in archive.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml') and '/' not in n[len('ppt/slides/'):]):
            node = ET.fromstring(archive.read(name))
            slides.append({'part': name, 'text': ''.join(t.text or '' for t in node.iter(f'{A}t'))})
        return {'part': part, 'properties': properties, 'slides': slides}


def inject(source, destination):
    with zipfile.ZipFile(source) as original:
        part = core_part(original)
        data = original.read(part)
        marker = b'</cp:coreProperties>'
        extra = (b'<cp:lastModifiedBy>Sentinel Maintainer</cp:lastModifiedBy>'
                 b'<cpTerms:created xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="cpTerms:W3CDTF">2020-01-02T03:04:05Z</cpTerms:created>'
                 b'<cpTerms:modified xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="cpTerms:W3CDTF">2021-02-03T04:05:06Z</cpTerms:modified>')
        if data.count(marker) != 1 or b'<cp:lastModifiedBy>' in data or b'<cpTerms:created' in data or b'<cpTerms:modified' in data:
            raise ValueError('Unexpected generated core properties layout')
        with zipfile.ZipFile(destination, 'w') as output:
            for entry in original.infolist():
                payload = original.read(entry.filename)
                output.writestr(entry, payload.replace(marker, extra + marker) if entry.filename == part else payload)


if __name__ == '__main__':
    if sys.argv[1] == 'inject':
        inject(sys.argv[2], sys.argv[3])
    elif sys.argv[1] == 'inspect':
        print(json.dumps(inventory(sys.argv[2]), ensure_ascii=False))
    else:
        raise SystemExit('Expected inject or inspect')
