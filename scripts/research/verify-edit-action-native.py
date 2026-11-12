#!/usr/bin/env python3
"""Independently check retained package identities and each hyperlink's owner and destination."""
import hashlib
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
RID = '{%s}id' % NS['r']
record = json.loads((ROOT / 'docs/testing/edit-action-powerpoint-evidence.json').read_text())
for item in record['artifacts'].values():
    assert hashlib.sha256((ROOT / item['path']).read_bytes()).hexdigest() == item['sha256']

def inspect(path):
    with zipfile.ZipFile(ROOT / path) as z:
        result = []
        for number in (1, 3):
            slide = ET.fromstring(z.read(f'ppt/slides/slide{number}.xml'))
            rels = ET.fromstring(z.read(f'ppt/slides/_rels/slide{number}.xml.rels'))
            by_id = {r.get('Id'): r for r in rels.findall('rel:Relationship', NS)}
            shapes = {shape.find('p:nvSpPr/p:cNvPr', NS).get('name'): shape for shape in slide.findall('.//p:sp', NS)}
            owners = [('shape', shapes['Go to third'].findall('p:nvSpPr/p:cNvPr/a:hlinkClick', NS))] if number == 1 else [
                ('text', shapes['run'].findall('.//a:rPr/a:hlinkClick', NS)),
                ('table', slide.findall('.//a:tbl//a:rPr/a:hlinkClick', NS))]
            for owner, links in owners:
                assert len(links) == 1
                link = links[0]
                rel = by_id[link.attrib[RID]]
                assert rel.get('Type') == NS['r'] + '/slide' and rel.get('TargetMode') is None
                assert link.get('action') == 'ppaction://hlinksldjump'
                result.append((owner, rel.get('Target'), link.get('tooltip')))
        return result

expected = [('shape', 'slide3.xml', 'Third slide'), ('text', 'slide1.xml', 'First slide'), ('table', 'slide2.xml', 'Second slide')]
assert inspect(record['artifacts']['source']['path']) == expected
assert inspect(record['artifacts']['saved']['path']) == expected
print('Retained action evidence verified.')
