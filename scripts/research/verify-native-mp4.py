#!/usr/bin/env python3
"""Independent inspection of the retained PowerPoint video sample."""
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
        assert len(pics) == 1
        pic = pics[0]
        object_id = pic.find('p:nvPicPr/p:cNvPr', NS).get('id')
        video = rels[pic.find('p:nvPicPr/p:nvPr/a:videoFile', NS).get(f"{{{NS['r']}}}link")]
        media = rels[pic.find('.//p14:media', NS).get(f"{{{NS['r']}}}embed")]
        poster = rels[pic.find('p:blipFill/a:blip', NS).get(f"{{{NS['r']}}}embed")]
        assert video.get('Type') == NS['r'] + '/video' and media.get('Type') == 'http://schemas.microsoft.com/office/2007/relationships/media'
        assert poster.get('Type') == NS['r'] + '/image'
        assert all(item.get('TargetMode', 'Internal') == 'Internal' for item in (video, media, poster))
        assert video.get('Target') == media.get('Target')
        video_part = posixpath.normpath(posixpath.join('ppt/slides', video.get('Target')))
        poster_part = posixpath.normpath(posixpath.join('ppt/slides', poster.get('Target')))
        assert video_part.startswith('ppt/media/') and poster_part.startswith('ppt/media/') and video_part != poster_part
        assert hashlib.sha256(z.read(video_part)).hexdigest() == EXPECTED
        assert z.read(poster_part).startswith(b'\x89PNG\r\n\x1a\n')
        types = ET.fromstring(z.read('[Content_Types].xml'))
        assert any(n.get('Extension') == 'mp4' and n.get('ContentType') == 'video/mp4' for n in types.findall('ct:Default', NS))
        timing = root.find('p:timing', NS)
        commands = timing.findall('.//p:cmd', NS)
        assert [c.get('cmd') for c in commands] == ['playFrom(0.0)', 'togglePause']
        assert all(c.find('.//p:spTgt', NS).get('spid') == object_id for c in commands)
        assert timing.find('.//p:video/p:cMediaNode/p:tgtEl/p:spTgt', NS).get('spid') == object_id
        assert commands[0].find('.//p:cTn', NS).get('dur') == '2000'
        ids = [n.get('id') for n in timing.findall('.//p:cTn', NS)]
        assert len(ids) == len(set(ids))
        return {'objectId': object_id, 'video': video.get('Target'), 'poster': poster.get('Target'), 'sha256': EXPECTED}


if __name__ == '__main__':
    print(verify(open(sys.argv[1], 'rb').read()))
