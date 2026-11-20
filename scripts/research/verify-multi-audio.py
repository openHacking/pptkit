#!/usr/bin/env python3
"""Independent ZIP/XML inspection of two same-slide audio references."""
import hashlib
import io
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main", "a": "http://schemas.openxmlformats.org/drawingml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main", "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
EXPECTED = "af6d351ebb501dff7a6b053de2b357ef5f873c4962738d7a7fc5c4cd19b98bdc"


def verify(data):
    with ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(z.read('ppt/slides/slide1.xml'))
        rels = {node.get('Id'): node for node in ET.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels')).findall('rel:Relationship', NS)}
        pics = root.findall('.//p:pic', NS)
        assert len(pics) == 2
        objects, paths = [], []
        for pic in pics:
            objects.append(pic.find('p:nvPicPr/p:cNvPr', NS).get('id'))
            audio = pic.find('p:nvPicPr/p:nvPr/a:audioFile', NS)
            media = pic.find('p:nvPicPr/p:nvPr/p:extLst/p:ext/p14:media', NS)
            first = rels[audio.get(f"{{{NS['r']}}}link")]
            second = rels[media.get(f"{{{NS['r']}}}embed")]
            assert first.get('Type') == NS['r'] + '/audio'
            assert second.get('Type') == 'http://schemas.microsoft.com/office/2007/relationships/media'
            assert first.get('Target') == second.get('Target')
            paths.append(first.get('Target'))
        assert len(set(objects)) == 2 and len(set(paths)) == 1
        media_parts = [name for name in z.namelist() if name.startswith('ppt/media/') and name.endswith('.wav')]
        assert len(media_parts) == 1 and hashlib.sha256(z.read(media_parts[0])).hexdigest() == EXPECTED
        assert paths[0] == '../media/' + media_parts[0].split('/')[-1]
        timing = root.find('p:timing', NS)
        commands = timing.findall('.//p:cmd', NS)
        assert [item.get('cmd') for item in commands] == ['playFrom(0.0)'] * 2
        assert [item.find('.//p:spTgt', NS).get('spid') for item in commands] == objects
        assert [item.find('p:cMediaNode/p:tgtEl/p:spTgt', NS).get('spid') for item in timing.findall('.//p:audio', NS)] == objects
        ids = [node.get('id') for node in timing.findall('.//p:cTn', NS)]
        assert len(ids) == len(set(ids))
        return {'objects': objects, 'wavPart': media_parts[0], 'timingIds': ids}


if __name__ == '__main__':
    print(verify(open(sys.argv[1], 'rb').read()))
