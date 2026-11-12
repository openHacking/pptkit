#!/usr/bin/env python3
"""Independent source slide order/visibility and headless PDF page inventory."""
import json
import posixpath
import sys
import zipfile
from xml.etree import ElementTree as ET

import fitz

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def inspect_pptx(path):
    with zipfile.ZipFile(path) as archive:
        if len(archive.namelist()) != len(set(archive.namelist())):
            raise ValueError('Duplicate ZIP entries')
        presentation = ET.fromstring(archive.read('ppt/presentation.xml'))
        rels = ET.fromstring(archive.read('ppt/_rels/presentation.xml.rels'))
        relationships = {}
        for rel in rels:
            identifier = rel.get('Id')
            if not identifier or identifier in relationships:
                raise ValueError('Missing or duplicate relationship identity')
            relationships[identifier] = rel
        slides = []
        for slide in presentation.findall(f'{P}sldIdLst/{P}sldId'):
            rel = relationships.get(slide.get(R + 'id'))
            if rel is None:
                raise ValueError('Slide relationship is missing')
            if rel.get('TargetMode') == 'External' or not rel.get('Type', '').endswith('/slide'):
                raise ValueError('Slide relationship is not internal')
            target = rel.get('Target')
            if not target or '\\' in target:
                raise ValueError('Slide relationship target is invalid')
            part = posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('ppt', target))
            if part.startswith('../') or part not in archive.namelist():
                raise ValueError(f'Invalid slide target: {target}')
            root = ET.fromstring(archive.read(part))
            raw = root.get('show')
            if raw is None or raw.strip() in ('1', 'true'):
                hidden = False
            elif raw.strip() in ('0', 'false'):
                hidden = True
            else:
                raise ValueError(f'Invalid slide show: {raw}')
            slides.append({'id': slide.get('id'), 'part': part, 'show': raw, 'hidden': hidden, 'text': ''.join(node.text or '' for node in root.iter(f'{A}t'))})
        return {'slides': slides}


def inspect_pdf(path):
    with fitz.open(path) as document:
        return {'pageCount': len(document), 'pages': [{'width': page.rect.width, 'height': page.rect.height, 'text': page.get_text()} for page in document]}


if __name__ == '__main__':
    print(json.dumps({'pptx': inspect_pptx(sys.argv[1]), 'pdf': inspect_pdf(sys.argv[2])}, ensure_ascii=False))
