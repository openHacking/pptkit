#!/usr/bin/env python3
"""Independent ZIP/XML checks for the two-video shared-resource sample."""
import hashlib
import io
import posixpath
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main', 'rel': 'http://schemas.openxmlformats.org/package/2006/relationships', 'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
EXPECTED = '6853306557b18631b94e76410bf4b51d41e10341467524a38d1e0fd29efa062a'


def verify(data):
    with ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(z.read('ppt/slides/slide1.xml'))
        rels = {r.get('Id'): r for r in ET.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels')).findall('rel:Relationship', NS)}
        pics = root.findall('.//p:pic', NS)
        assert len(pics) == 2
        objects, targets = [], []
        for pic in pics:
            objects.append(pic.find('p:nvPicPr/p:cNvPr', NS).get('id'))
            direct = rels[pic.find('p:nvPicPr/p:nvPr/a:videoFile', NS).get(f"{{{NS['r']}}}link")]
            media = rels[pic.find('.//p14:media', NS).get(f"{{{NS['r']}}}embed")]
            poster = rels[pic.find('p:blipFill/a:blip', NS).get(f"{{{NS['r']}}}embed")]
            assert direct.get('Type') == NS['r'] + '/video' and media.get('Type') == 'http://schemas.microsoft.com/office/2007/relationships/media' and poster.get('Type') == NS['r'] + '/image'
            assert all(r.get('TargetMode', 'Internal') == 'Internal' for r in (direct, media, poster))
            assert direct.get('Target') == media.get('Target')
            target = posixpath.normpath(posixpath.join('ppt/slides', direct.get('Target')))
            image = posixpath.normpath(posixpath.join('ppt/slides', poster.get('Target')))
            assert target.startswith('ppt/media/') and image.startswith('ppt/media/') and image != target
            assert hashlib.sha256(z.read(target)).hexdigest() == EXPECTED and z.read(image).startswith(b'\x89PNG\r\n\x1a\n')
            targets.append(target)
        assert len(set(objects)) == 2 and len(set(targets)) == 1
        assert [name for name in z.namelist() if name.startswith('ppt/media/') and name.endswith('.mp4')] == [targets[0]]
        types = ET.fromstring(z.read('[Content_Types].xml'))
        assert any(n.get('Extension') == 'mp4' and n.get('ContentType') == 'video/mp4' for n in types.findall('ct:Default', NS))
        timing = root.find('p:timing', NS)
        cmds = timing.findall('.//p:cmd', NS)
        assert [c.get('cmd') for c in cmds] == ['playFrom(0.0)'] * 2 + ['togglePause'] * 2
        assert [c.find('.//p:spTgt', NS).get('spid') for c in cmds] == objects * 2
        assert [c.find('.//p:cTn', NS).get('dur') for c in cmds] == ['2000', '2000', '1', '1']
        assert [v.find('p:cMediaNode/p:tgtEl/p:spTgt', NS).get('spid') for v in timing.findall('.//p:video', NS)] == objects
        interactive = [n for n in timing.findall('.//p:seq', NS) if n.find('p:cTn', NS).get('nodeType') == 'interactiveSeq']
        assert [n.find('p:cTn/p:stCondLst/p:cond/p:tgtEl/p:spTgt', NS).get('spid') for n in interactive] == objects
        assert [n.find('p:nextCondLst/p:cond/p:tgtEl/p:spTgt', NS).get('spid') for n in interactive] == objects
        ids = [n.get('id') for n in timing.findall('.//p:cTn', NS)]
        assert len(ids) == len(set(ids))
        return {'objects': objects, 'video': targets[0], 'timingIds': ids}


if __name__ == '__main__':
    print(verify(open(sys.argv[1], 'rb').read()))
