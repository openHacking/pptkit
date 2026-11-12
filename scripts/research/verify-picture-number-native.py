#!/usr/bin/env python3
"""Check the named picture/number fixture after a native save and reordering."""
import hashlib
import json
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as E

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def relationship_map(archive, part):
    rels = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
    root = E.fromstring(archive.read(rels))
    result = {}
    for item in root:
        key = item.get('Id')
        if not key or key in result or item.get('TargetMode') == 'External':
            raise ValueError('Invalid or external relationship')
        result[key] = (item.get('Type'), posixpath.normpath(posixpath.join(posixpath.dirname(part), item.get('Target'))))
    return result


def inspect(path):
    with zipfile.ZipFile(path) as archive:
        if len(archive.namelist()) != len(set(archive.namelist())): raise ValueError('Duplicate ZIP member')
        main = 'ppt/presentation.xml'
        presentation = E.fromstring(archive.read(main))
        if presentation.get('firstSlideNum') != '17': raise ValueError('Wrong starting number')
        rels = relationship_map(archive, main)
        ids = presentation.findall('./' + P + 'sldIdLst/' + P + 'sldId')
        if len(ids) != 2: raise ValueError('Wrong slide count')
        pages = []
        for item in ids:
            kind, name = rels[item.get(R + 'id')]
            if not kind.endswith('/slide'): raise ValueError('Wrong slide relationship')
            root = E.fromstring(archive.read(name))
            pictures = {}
            slide_rels = relationship_map(archive, name)
            for pic in root.iter(P + 'pic'):
                identity = pic.find('./' + P + 'nvPicPr/' + P + 'cNvPr').get('name')
                if identity in pictures: raise ValueError('Duplicate picture identity')
                blip = pic.find('./' + P + 'blipFill/' + A + 'blip')
                kind, media = slide_rels[blip.get(R + 'embed')]
                if not kind.endswith('/image'): raise ValueError('Wrong image relationship')
                crop = pic.find('./' + P + 'blipFill/' + A + 'srcRect')
                rect = {k: int(crop.get(k, '0')) for k in ('l', 't', 'r', 'b')}
                transform = pic.find('./' + P + 'spPr/' + A + 'xfrm')
                geometry = tuple(int(transform.find(A + tag).get(attr)) for tag, attr in [('off','x'),('off','y'),('ext','cx'),('ext','cy')])
                pictures[identity] = {'sha256': hashlib.sha256(archive.read(media)).hexdigest(), 'crop': rect, 'geometry': geometry}
            pages.append({'pictures': pictures, 'text': ''.join(n.text or '' for n in root.iter(A + 't'))})
        return pages


def verify(candidate, native, position):
    original, actual = inspect(candidate), inspect(native)
    if position not in (0, 1): raise ValueError('Invalid position')
    expected = original[0]['pictures']
    if set(expected) != {'target', 'sibling'} or original[1]['pictures'] or original[1]['text'] != 'Control page': raise ValueError('Unexpected candidate')
    if expected['target']['sha256'] == expected['sibling']['sha256']: raise ValueError('Indistinguishable images')
    if any(p['crop'] != {'l': 12500, 't': 0, 'r': 0, 'b': 0} for p in expected.values()): raise ValueError('Unexpected candidate crop')
    if actual[position]['pictures'] != expected or actual[1-position]['pictures'] or actual[1-position]['text'] != 'Control page': raise ValueError('Incorrect native page, picture, crop, geometry or media')
    return {'picture_page': position + 1, 'starting_number': 17, 'pictures': expected}


if __name__ == '__main__':
    print(json.dumps(verify(sys.argv[1], sys.argv[2], int(sys.argv[3])), sort_keys=True))
