#!/usr/bin/env python3
"""Independently compare three named PowerPoint media-start samples."""
import hashlib
import io
import posixpath
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
HASHES = {'video': '6853306557b18631b94e76410bf4b51d41e10341467524a38d1e0fd29efa062a', 'audio': 'af6d351ebb501dff7a6b053de2b357ef5f873c4962738d7a7fc5c4cd19b98bdc'}


def inspect(data):
    with ZipFile(io.BytesIO(data)) as package:
        root = ET.fromstring(package.read('ppt/slides/slide1.xml'))
        rels = {r.get('Id'): r for r in ET.fromstring(package.read('ppt/slides/_rels/slide1.xml.rels')).findall('rel:Relationship', NS)}
        pictures = root.findall('.//p:pic', NS)
        assert len(pictures) == 4
        for picture in pictures:
            obj = picture.find('p:nvPicPr/p:cNvPr', NS).get('id')
            kind = 'video' if obj in ('2', '3') else 'audio'
            assert obj in ('2', '3', '4', '5')
            declaration = picture.find('p:nvPicPr/p:nvPr/a:' + kind + 'File', NS)
            relation = rels[declaration.get(f"{{{NS['r']}}}link")]
            assert relation.get('Type') == NS['r'] + '/' + kind and relation.get('TargetMode', 'Internal') == 'Internal'
            target = posixpath.normpath(posixpath.join('ppt/slides', relation.get('Target')))
            assert hashlib.sha256(package.read(target)).hexdigest() == HASHES[kind]
    timing = root.find('p:timing', NS)
    parents = {child: parent for parent in timing.iter() for child in parent}
    commands = timing.findall('.//p:cmd', NS)
    plays = [x for x in commands if x.get('cmd') == 'playFrom(0.0)']
    assert len(plays) == 4
    targets = [x.find('.//p:spTgt', NS).get('spid') for x in plays]
    assert sorted(targets) == ['2', '3', '4', '5']
    assert [x.find('.//p:spTgt', NS).get('spid') for x in commands if x.get('cmd') == 'togglePause'] == ['2', '3']
    ids = [x.get('id') for x in timing.findall('.//p:cTn', NS)]
    assert len(ids) == len(set(ids))
    def ancestry(command):
        path = []; node = command
        while node in parents:
            node = parents[node]; path.append(node)
        return path
    def effect(target):
        cmd = plays[targets.index(target)]
        chain = ancestry(cmd)
        effect_node = next(x for x in chain if x.tag == f"{{{NS['p']}}}cTn" and x.get('presetClass') == 'mediacall')
        sequence = next(x for x in chain if x.tag == f"{{{NS['p']}}}seq")
        containing = parents[effect_node]
        wrapper = parents[parents[containing]]
        return cmd, effect_node, wrapper, sequence, chain
    _, _, _, first_seq, first_chain = effect('4')
    cmd, node, wrapper, sequence, chain = effect('5')
    assert cmd.find('p:cBhvr/p:cTn', NS).get('dur') == '1000'
    assert all(effect(target)[1].get('nodeType') == 'clickEffect' and effect(target)[3].find('p:cTn', NS).get('nodeType') == 'mainSeq' for target in ['2', '3', '4'])
    assert effect('4')[0].find('p:cBhvr/p:cTn', NS).get('dur') == '1000'
    kind = sequence.find('p:cTn', NS).get('nodeType')
    delay = wrapper.find('p:stCondLst/p:cond', NS).get('delay')
    shared = any(x is y and x.tag == f"{{{NS['p']}}}par" and x in list(sequence.find('p:cTn/p:childTnLst', NS)) for x in first_chain for y in chain)
    trigger = sequence.find('p:cTn/p:stCondLst/p:cond', NS)
    return {'sequence': kind, 'effect': node.get('nodeType'), 'delay': delay, 'sameGroupAsPrevious': shared, 'trigger': trigger.get('evt') if trigger is not None else None, 'triggerTarget': trigger.find('p:tgtEl/p:spTgt', NS).get('spid') if trigger is not None and trigger.find('p:tgtEl/p:spTgt', NS) is not None else None}


def verify(baseline, automatic, object_click):
    assert inspect(baseline) == {'sequence': 'mainSeq', 'effect': 'clickEffect', 'delay': '0', 'sameGroupAsPrevious': False, 'trigger': None, 'triggerTarget': None}
    assert inspect(automatic) == {'sequence': 'mainSeq', 'effect': 'afterEffect', 'delay': '1000', 'sameGroupAsPrevious': True, 'trigger': None, 'triggerTarget': None}
    assert inspect(object_click) == {'sequence': 'interactiveSeq', 'effect': 'clickEffect', 'delay': '0', 'sameGroupAsPrevious': False, 'trigger': 'onClick', 'triggerTarget': '5'}


def verify_video_auto(data):
    """Verify the specific PowerPoint menu edit, without inferring slideshow playback."""
    with ZipFile(io.BytesIO(data)) as package:
        root = ET.fromstring(package.read('ppt/slides/slide1.xml'))
        rels = {r.get('Id'): r for r in ET.fromstring(package.read('ppt/slides/_rels/slide1.xml.rels')).findall('rel:Relationship', NS)}
        pictures = root.findall('.//p:pic', NS)
        assert [picture.find('p:nvPicPr/p:cNvPr', NS).get('id') for picture in pictures] == ['2', '3', '4', '5']
        for picture, kind in zip(pictures, ['video', 'video', 'audio', 'audio']):
            node = picture.find('p:nvPicPr/p:nvPr/a:' + kind + 'File', NS)
            relation = rels[node.get(f"{{{NS['r']}}}link")]
            assert relation.get('Type') == NS['r'] + '/' + kind and relation.get('TargetMode', 'Internal') == 'Internal'
            target = posixpath.normpath(posixpath.join('ppt/slides', relation.get('Target')))
            assert hashlib.sha256(package.read(target)).hexdigest() == HASHES[kind]
    timing = root.find('p:timing', NS)
    ids = [x.get('id') for x in timing.findall('.//p:cTn', NS)]
    assert all(x is not None and x.isdigit() for x in ids) and len({int(x) for x in ids}) == len(ids)
    main = timing.findall('.//p:cTn[@nodeType="mainSeq"]', NS)
    assert len(main) == 1
    groups = main[0].findall('p:childTnLst/p:par/p:cTn', NS)
    actual = []
    for group in groups:
        assert [c.attrib for c in group.findall('p:stCondLst/p:cond', NS)] == [{'delay': 'indefinite'}]
        row = []
        for wrapper in group.findall('p:childTnLst/p:par/p:cTn', NS):
            delay = wrapper.find('p:stCondLst/p:cond', NS).get('delay')
            effects = [x for x in wrapper.findall('.//p:cTn', NS) if x.get('presetClass') == 'mediacall']
            assert len(effects) == 1
            effect = effects[0]; command = effect.find('p:childTnLst/p:cmd', NS)
            assert command.get('cmd') == 'playFrom(0.0)'
            target = command.find('p:cBhvr/p:tgtEl/p:spTgt', NS).get('spid')
            duration = command.find('p:cBhvr/p:cTn', NS).get('dur')
            row.append((target, effect.get('nodeType'), delay, duration))
        actual.append(row)
    assert actual == [[('3', 'clickEffect', '0', '2000')], [('4', 'clickEffect', '0', '1000'), ('5', 'afterEffect', '1000', '1000'), ('2', 'afterEffect', '2000', '2000')]]
    pauses = timing.findall('.//p:cmd[@cmd="togglePause"]', NS)
    assert sorted(x.find('p:cBhvr/p:tgtEl/p:spTgt', NS).get('spid') for x in pauses) == ['2', '3']
    assert len(timing.findall('.//p:cmd', NS)) == 6
    return actual


if __name__ == '__main__':
    verify(*(open(path, 'rb').read() for path in sys.argv[1:]))
    print('Verified native sequence-click, automatic and object-click timing')
