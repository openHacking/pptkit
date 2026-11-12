#!/usr/bin/env python3
"""Inspect one PowerPoint image-replacement observation; does not drive the UI."""
import hashlib
import json
from pathlib import Path
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

def sha(data):
    return hashlib.sha256(data).hexdigest()

def pictures(path):
    with zipfile.ZipFile(path) as archive:
        slide = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
        rels = ET.fromstring(archive.read('ppt/slides/_rels/slide1.xml.rels'))
        relationships = {x.get('Id'): x for x in rels.findall(f'{{{REL}}}Relationship')}
        result = {}
        for pic in slide.findall('.//p:pic', NS):
            name = pic.find('p:nvPicPr/p:cNvPr', NS).get('name')
            if name in result: raise ValueError('duplicate picture name')
            rid = pic.find('p:blipFill/a:blip', NS).get(f'{{{R}}}embed')
            rel = relationships.get(rid)
            if rel is None or not rel.get('Type', '').endswith('/image') or rel.get('TargetMode') == 'External': raise ValueError('invalid picture relationship')
            part = posixpath.normpath(posixpath.join('ppt/slides', rel.get('Target', '')))
            crop = pic.find('p:blipFill/a:srcRect', NS)
            transform = pic.find('p:spPr/a:xfrm', NS)
            off, ext = transform.find('a:off', NS), transform.find('a:ext', NS)
            result[name] = {'sha256': sha(archive.read(part)), 'part': part,
                            'crop': None if crop is None else {k: int(crop.get(k, '0')) for k in ('l', 't', 'r', 'b')},
                            'frame': [int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy'))]}
        return result

def inspect(source, output, replacement):
    before, after = pictures(source), pictures(output)
    names = {'image-stretch', 'image-contain', 'image-cover', 'image-crop'}
    if set(before) != names or set(after) != names: raise ValueError('picture inventory differs')
    original = before['image-stretch']['sha256']
    if any(item['sha256'] != original for item in before.values()): raise ValueError('source media differs')
    if after['image-stretch']['sha256'] != sha(replacement.read_bytes()) or original == after['image-stretch']['sha256']: raise ValueError('replacement media differs')
    if before['image-stretch']['frame'] != [508000, 1016000, 2286000, 1397000] or after['image-stretch']['frame'] != [533400, 1016000, 2235200, 1397000]: raise ValueError('replacement frame differs')
    for name in names - {'image-stretch'}:
        if any(before[name][key] != after[name][key] for key in ('sha256', 'crop', 'frame')): raise ValueError(f'control changed: {name}')
    return {'source': before, 'saved': after}

if __name__ == '__main__':
    if sys.argv[1:] not in ([], ['--local']): raise SystemExit('usage: verify-corpus-image-replacement.py [--local]')
    evidence = json.loads((ROOT / 'docs/testing/corpus-image-replacement-powerpoint-evidence.json').read_text())
    for item in evidence['artifacts'].values():
        if sha((ROOT / item['path']).read_bytes()) != item['sha256']: raise ValueError('artifact hash differs')
    artifacts = evidence['artifacts']
    inspect(ROOT / artifacts['source']['path'], ROOT / artifacts['output']['path'], ROOT / artifacts['replacement']['path'])
    if sys.argv[1:]:
        directory = ROOT / evidence['artifactDirectory']
        manifest = json.loads((directory / 'manifest.json').read_text())
        fixture = next(item for item in manifest['fixtures'] if item['id'] == 'images')
        if fixture['sha256'] != artifacts['source']['sha256'] or sha((directory / 'images.pptx').read_bytes()) != artifacts['source']['sha256']: raise ValueError('source manifest differs')
        if sha((directory / 'images-replacement-edit.pptx').read_bytes()) != artifacts['output']['sha256']: raise ValueError('local output differs')
        for name, expected in evidence['captures'].items():
            if sha((directory / name).read_bytes()) != expected: raise ValueError('capture hash differs')
        for name in ('images-replacement-first.ax.txt', 'images-replacement-reopen.ax.txt'):
            state = (directory / name).read_text()
            if any(label not in state for label in ('images-replacement-edit', 'image-stretch', 'image-contain', 'image-cover', 'image-crop')): raise ValueError('captured state differs')
    print('verified image replacement package' + (' and local captures' if sys.argv[1:] else ''))
