#!/usr/bin/env python3
"""Independent XML/ZIP oracle for the named shared-picture candidate."""
import io
import json
import sys
import zipfile
import xml.etree.ElementTree as E

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
MAIN = 'ppt/presentation.xml'; SLIDE = 'ppt/slides/slide1.xml'; RELS = 'ppt/slides/_rels/slide1.xml.rels'; TYPES = '[Content_Types].xml'; MEDIA = 'ppt/media/pptkit-image-1.png'


def parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)): raise ValueError('Duplicate ZIP member')
        return {n: archive.read(n) for n in names}


def same(before, after, edit):
    a, b = E.fromstring(before), E.fromstring(after); edit(a)
    if E.tostring(a) != E.tostring(b): raise ValueError('Unexpected XML change')


def verify(source, candidate, replacement):
    old, new = parts(source), parts(candidate)
    if set(new) - set(old) != {MEDIA} or set(old) - set(new): raise ValueError('Unexpected inventory')
    if new[MEDIA] != replacement: raise ValueError('Incorrect replacement media')
    changed = {name for name in old if old[name] != new[name]}
    if changed != {MAIN, SLIDE, RELS, TYPES}: raise ValueError('Unexpected changed parts')
    def number(root):
        if root.tag != P + 'presentation' or 'firstSlideNum' in root.attrib: raise ValueError('Unexpected source numbering')
        root.set('firstSlideNum', '17')
    def slide(root):
        pictures = list(root.iter(P + 'pic'))
        if len(pictures) != 2 or [n.find('.//' + P + 'cNvPr').get('name') for n in pictures] != ['target', 'sibling']: raise ValueError('Unexpected picture identities')
        blips = [n.find('.//' + A + 'blip') for n in pictures]
        if [b.get(R + 'embed') for b in blips] != ['rId3', 'rId4']: raise ValueError('Unexpected source embeds')
        if any(n.find('.//' + A + 'srcRect').get('l') != '12500' for n in pictures): raise ValueError('Unexpected crop')
        blips[0].set(R + 'embed', 'pptkitImage1')
    def relationships(root):
        ids = [n.get('Id') for n in root]
        if len(ids) != len(set(ids)) or 'pptkitImage1' in ids: raise ValueError('Unexpected relationship IDs')
        for rid in ('rId3', 'rId4'):
            matching = [n for n in root if n.get('Id') == rid]
            if len(matching) != 1 or matching[0].attrib != {'Id': rid, 'Type': R[1:-1] + '/image', 'Target': '../media/shared.png'}:
                raise ValueError('Source pictures do not share the expected embedded image')
        if 'ppt/media/shared.png' not in old or old['ppt/media/shared.png'] == replacement:
            raise ValueError('Missing or indistinguishable source image')
        E.SubElement(root, PKG + 'Relationship', {'Id': 'pptkitImage1', 'Type': R[1:-1] + '/image', 'Target': '../../ppt/media/pptkit-image-1.png'})
    def types(root):
        E.SubElement(root, CT + 'Override', {'PartName': '/' + MEDIA, 'ContentType': 'image/png'})
    for part, edit in [(MAIN, number), (SLIDE, slide), (RELS, relationships), (TYPES, types)]: same(old[part], new[part], edit)
    return {'starting_number': 17, 'picture_count': 2, 'target_media': MEDIA, 'crop_left': 12500, 'changed_parts': sorted(changed)}


if __name__ == '__main__':
    with open(sys.argv[1], 'rb') as f: source = f.read()
    with open(sys.argv[2], 'rb') as f: candidate = f.read()
    with open(sys.argv[3], 'rb') as f: replacement = f.read()
    print(json.dumps(verify(source, candidate, replacement), sort_keys=True))
