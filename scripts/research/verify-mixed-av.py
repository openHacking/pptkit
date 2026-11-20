#!/usr/bin/env python3
"""Independent ZIP/XML verification of two video and two audio instances."""
import hashlib
import io
import posixpath
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main', 'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
HASHES = {'video': '6853306557b18631b94e76410bf4b51d41e10341467524a38d1e0fd29efa062a', 'audio': 'af6d351ebb501dff7a6b053de2b357ef5f873c4962738d7a7fc5c4cd19b98bdc'}


def verify(data):
    with ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(z.read('ppt/slides/slide1.xml'))
        rels = {r.get('Id'): r for r in ET.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels')).findall('rel:Relationship', NS)}
        pics = root.findall('.//p:pic', NS)
        assert len(pics) == 4
        objects, by_kind, targets = [], {'audio': [], 'video': []}, {'audio': set(), 'video': set()}
        for pic in pics:
            obj = pic.find('p:nvPicPr/p:cNvPr', NS).get('id')
            nv = pic.find('p:nvPicPr/p:nvPr', NS)
            kind = 'video' if nv.find('a:videoFile', NS) is not None else 'audio'
            node = nv.find('a:videoFile' if kind == 'video' else 'a:audioFile', NS)
            direct = rels[node.get(f"{{{NS['r']}}}link")]
            media = rels[nv.find('.//p14:media', NS).get(f"{{{NS['r']}}}embed")]
            poster = rels[pic.find('p:blipFill/a:blip', NS).get(f"{{{NS['r']}}}embed")]
            assert direct.get('Type') == NS['r'] + '/' + kind and media.get('Type') == 'http://schemas.microsoft.com/office/2007/relationships/media' and poster.get('Type') == NS['r'] + '/image'
            assert all(r.get('TargetMode', 'Internal') == 'Internal' for r in (direct, media, poster))
            assert direct.get('Target') == media.get('Target')
            target = posixpath.normpath(posixpath.join('ppt/slides', direct.get('Target')))
            image = posixpath.normpath(posixpath.join('ppt/slides', poster.get('Target')))
            assert target.startswith('ppt/media/') and image.startswith('ppt/media/') and target != image
            assert hashlib.sha256(z.read(target)).hexdigest() == HASHES[kind]
            assert z.read(image).startswith(b'\x89PNG\r\n\x1a\n')
            objects.append(obj); by_kind[kind].append(obj); targets[kind].add(target)
        assert len(set(objects)) == 4 and all(len(targets[kind]) == 1 and len(by_kind[kind]) == 2 for kind in targets)
        timing = root.find('p:timing', NS)
        commands = timing.findall('.//p:cmd', NS)
        assert [c.get('cmd') for c in commands] == ['playFrom(0.0)'] * 4 + ['togglePause'] * 2
        assert [c.find('.//p:spTgt', NS).get('spid') for c in commands] == objects + by_kind['video']
        assert [c.find('.//p:cTn', NS).get('dur') for c in commands] == ['2000' if x in by_kind['video'] else '1000' for x in objects] + ['1', '1']
        for kind in by_kind:
            assert [n.find('p:cMediaNode/p:tgtEl/p:spTgt', NS).get('spid') for n in timing.findall('.//p:' + kind, NS)] == by_kind[kind]
        sequences = [n for n in timing.findall('.//p:seq', NS) if n.find('p:cTn', NS).get('nodeType') == 'interactiveSeq']
        assert [n.find('p:cTn/p:stCondLst/p:cond/p:tgtEl/p:spTgt', NS).get('spid') for n in sequences] == by_kind['video']
        assert [n.find('p:nextCondLst/p:cond/p:tgtEl/p:spTgt', NS).get('spid') for n in sequences] == by_kind['video']
        ids = [n.get('id') for n in timing.findall('.//p:cTn', NS)]
        assert len(ids) == len(set(ids))
        return {'objects': objects, 'byKind': by_kind, 'timingIds': ids}


if __name__ == '__main__':
    print(verify(open(sys.argv[1], 'rb').read()))
