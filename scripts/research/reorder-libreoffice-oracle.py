#!/usr/bin/env python3
import copy
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import fitz

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
KINDS = {f'{P}sp', f'{P}pic', f'{P}cxnSp', f'{P}grpSp', f'{P}graphicFrame'}


def inventory(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('ppt/slides/slide1.xml'))
        tree = root.find(f'.//{P}spTree')
        def node(element):
            nonvisual = element.find(f'{P}nvGrpSpPr/{P}cNvPr') if element.tag == f'{P}grpSp' else element.find(f'{P}nvSpPr/{P}cNvPr')
            transform = element.find(f'{P}grpSpPr/{A}xfrm') if element.tag == f'{P}grpSp' else element.find(f'{P}spPr/{A}xfrm')
            geometry = None if transform is None else {'attributes': dict(transform.attrib), **{child.tag.split('}')[-1]: dict(child.attrib) for child in transform if child.tag in {f'{A}off', f'{A}ext', f'{A}chOff', f'{A}chExt'}}}
            detached = copy.deepcopy(element)
            for child in list(detached):
                if child.tag in KINDS:
                    detached.remove(child)
            own_nonvisual = detached.find(f'{P}nvGrpSpPr/{P}cNvPr') if element.tag == f'{P}grpSp' else detached.find(f'{P}nvSpPr/{P}cNvPr')
            if own_nonvisual is not None:
                own_nonvisual.attrib.pop('id', None)
            return {'kind': element.tag.split('}')[-1], 'name': None if nonvisual is None else nonvisual.get('name'), 'description': None if nonvisual is None else nonvisual.get('descr'), 'transform': geometry, 'propertiesSha256': hashlib.sha256(ET.tostring(detached)).hexdigest(), 'children': [node(child) for child in element if child.tag in KINDS]}
        objects = [node(element) for element in tree if element.tag in KINDS]
        text = ''.join(node.text or '' for node in root.iter(f'{A}t'))
        notes = ''.join(node.text or '' for node in ET.fromstring(z.read('ppt/notesSlides/notesSlide1.xml')).iter(f'{A}t'))
        return {'objects': objects, 'text': text, 'notes': notes}


def raster(path):
    with fitz.open(path) as doc:
        if len(doc) != 1:
            raise ValueError(f'Expected one page: {path}')
        page = doc[0]
        pix = page.get_pixmap(alpha=False)
        return {'width': pix.width, 'height': pix.height, 'components': pix.n, 'samples': pix.samples, 'text': page.get_text()}


def difference(a, b):
    if (a['width'], a['height'], a['components']) != (b['width'], b['height'], b['components']):
        return {'status': 'size-mismatch'}
    channels = a['components']
    return {'status': 'compared', 'changedPixels': sum(a['samples'][i:i+channels] != b['samples'][i:i+channels] for i in range(0, len(a['samples']), channels)), 'totalPixels': a['width'] * a['height']}


def check_source(items):
    assert [x['name'] for x in items[0]['objects']] == ['A', 'B', 'Outer', 'Caption']
    assert [x['name'] for x in items[1]['objects']] == ['B', 'A', 'Outer', 'Caption']
    for item, expected_outer, expected_inner in [(items[0], ['C', 'D', 'Inner'], ['E', 'F']), (items[1], ['D', 'C', 'Inner'], ['F', 'E'])]:
        group = item['objects'][2]
        assert [x['name'] for x in group['children']] == expected_outer
        assert [x['name'] for x in group['children'][2]['children']] == expected_inner


def identity_tree(nodes):
    return [{'kind': node['kind'], 'name': node['name'], 'children': identity_tree(node['children'])} for node in nodes]


def named_transforms(nodes):
    result = {}
    for node in nodes:
        if node['name'] in result:
            raise ValueError(f"Duplicate name: {node['name']}")
        result[node['name']] = node['transform']
        for name, transform in named_transforms(node['children']).items():
            if name in result:
                raise ValueError(f'Duplicate name: {name}')
            result[name] = transform
    return result


def main(directory):
    data = []
    with zipfile.ZipFile(directory / 'input/control.pptx') as source, zipfile.ZipFile(directory / 'input/reordered.pptx') as reordered:
        if set(source.namelist()) != set(reordered.namelist()):
            raise ValueError('Source package entry sets differ')
        different_parts = [name for name in source.namelist() if source.read(name) != reordered.read(name)]
        if different_parts != ['ppt/slides/slide1.xml']:
            raise ValueError(f'Unexpected source payload differences: {different_parts}')
    for round in range(4):
        folder = directory / ('input' if round == 0 else f'round-{round}')
        items = [inventory(folder / f'{name}.pptx') for name in ('control', 'reordered')]
        if round == 0:
            check_source(items)
            def properties(nodes):
                return {entry['name']: entry['propertiesSha256'] for entry in nodes} | {name: value for entry in nodes for name, value in properties(entry['children']).items()}
            if properties(items[0]['objects']) != properties(items[1]['objects']):
                raise ValueError('Source object properties differ beyond their drawing IDs and order')
        for item in items:
            if 'Reorder control text' not in item['text'] or 'Reorder control notes' not in item['notes']:
                raise ValueError('Text or notes missing')
        images = [raster(directory / f'pdf-{round}' / f'{name}.pdf') for name in ('control', 'reordered')]
        for image in images:
            if 'Reordercontroltext' not in ''.join(image['text'].split()):
                raise ValueError('PDF text missing')
        data.append({'round': round, 'control': items[0], 'reordered': items[1], 'withinRound': difference(*images), 'pdfText': [image['text'] for image in images]})
    base = [raster(directory / 'pdf-0' / f'{name}.pdf') for name in ('control', 'reordered')]
    for item in data:
        images = [raster(directory / f"pdf-{item['round']}" / f'{name}.pdf') for name in ('control', 'reordered')]
        item['controlVsInitial'] = difference(base[0], images[0])
        item['reorderedVsInitial'] = difference(base[1], images[1])
        for key in ('control', 'reordered'):
            item[f'{key}IdentityTreeMatchesInitial'] = identity_tree(item[key]['objects']) == identity_tree(data[0][key]['objects'])
            item[f'{key}InventoryMatchesInitial'] = item[key]['objects'] == data[0][key]['objects']
    top_level_extents = {key: {name: [named_transforms(item[key]['objects'])[name]['ext'] for item in data] for name in ('A', 'B')} for key in ('control', 'reordered')}
    print(json.dumps({'method': 'Python zipfile/ElementTree and PyMuPDF; source order and non-order properties asserted, subsequent producer orders observed', 'sourceDifferentParts': different_parts, 'topLevelExtentsByRound': top_level_extents, 'rounds': data}))


if __name__ == '__main__':
    main(Path(sys.argv[1]))
