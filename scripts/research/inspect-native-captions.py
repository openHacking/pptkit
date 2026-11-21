#!/usr/bin/env python3
"""Research-only inventory of observed PowerPoint caption tracks; not a VTT validator."""
import json
import posixpath
import sys
from collections import Counter
from xml.etree import ElementTree as ET
from zipfile import ZipFile

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
P14 = '{http://schemas.microsoft.com/office/powerpoint/2010/main}'
P173 = '{http://schemas.microsoft.com/office/powerpoint/2017/3/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
TRACK_REL = 'http://schemas.microsoft.com/office/2017/04/relationships/track'


def inventory(path):
    with ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError('duplicate archive names')
        types = ET.fromstring(archive.read('[Content_Types].xml'))
        defaults = {x.get('Extension', '').lower(): x.get('ContentType') for x in types.findall(CT + 'Default')}
        overrides = {x.get('PartName', '').lstrip('/'): x.get('ContentType') for x in types.findall(CT + 'Override')}
        result = []
        slides = sorted(x for x in names if x.startswith('ppt/slides/slide') and x.endswith('.xml') and '/' not in x[len('ppt/slides/'):])
        for slide in slides:
            root = ET.fromstring(archive.read(slide))
            rel_path = posixpath.join(posixpath.dirname(slide), '_rels', posixpath.basename(slide) + '.rels')
            rel_root = ET.fromstring(archive.read(rel_path)) if rel_path in names else ET.Element('missing')
            relations = rel_root.findall(REL + 'Relationship')
            counts = Counter(x.get('Id') for x in relations)
            rels = {x.get('Id'): x for x in relations if counts[x.get('Id')] == 1}
            for pic in root.iter(P + 'pic'):
                ident = pic.find('./' + P + 'nvPicPr/' + P + 'cNvPr')
                # Include only pictures with an observed media extension, not ordinary images.
                media = list(pic.iter(P14 + 'media'))
                if ident is None or not media:
                    continue
                tracks = []
                for node in pic.iter(P173 + 'track'):
                    rid = node.get(R + 'embed')
                    relation = rels.get(rid)
                    record = {'id': node.get('id'), 'language': node.get('lang'), 'label': node.get('label'), 'relationshipId': rid}
                    if relation is None or relation.get('Type') != TRACK_REL:
                        record['status'] = 'invalid-relationship'
                    elif relation.get('TargetMode', 'Internal') != 'Internal':
                        record['status'] = 'external-not-loaded'
                    else:
                        target = relation.get('Target', '')
                        part = posixpath.normpath(posixpath.join(posixpath.dirname(slide), target))
                        if not target or target.startswith('/') or '#' in target or '?' in target or part.startswith('../') or part not in names:
                            record['status'] = 'missing-or-invalid-target'
                        else:
                            record['part'] = part
                            record['contentType'] = overrides.get(part, defaults.get(part.rsplit('.', 1)[-1].lower()))
                            if record['contentType'] != 'text/vtt':
                                record['status'] = 'unsupported-content-type'
                            else:
                                try:
                                    payload = archive.read(part).decode('utf-8-sig')
                                    if not (payload.startswith('WEBVTT\n') or payload.startswith('WEBVTT\r\n')):
                                        record['status'] = 'invalid-vtt-header'
                                    else:
                                        record['status'] = 'vtt-header-only'
                                        record['text'] = payload
                                except UnicodeDecodeError:
                                    record['status'] = 'invalid-utf8'
                    tracks.append(record)
                trims = [dict(x.attrib) for x in pic.iter(P14 + 'trim')]
                infos = [dict(x.attrib) for x in pic.iter(P173 + 'tracksInfo')]
                result.append({'slide': slide, 'objectId': ident.get('id'), 'trims': trims, 'trackInfo': infos, 'tracks': tracks})
        return result


if __name__ == '__main__':
    print(json.dumps(inventory(sys.argv[1]), ensure_ascii=False, sort_keys=True))
