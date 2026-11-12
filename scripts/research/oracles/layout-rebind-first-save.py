"""Strict first-save comparison for the named three-page layout-rebind control."""
import copy
import importlib.util
import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as E
import fitz

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('rebind_chart_review', HERE/'master-chart-first-save.py')
h = importlib.util.module_from_spec(spec)
spec.loader.exec_module(h)
a, n = h.a, h.a.n
A, P, R, NS = h.A, h.c.P, h.c.R, h.c.NS
A16 = 'http://schemas.microsoft.com/office/drawing/2014/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
DCT = 'http://purl.org/dc/terms/'


def identity(parts):
    labels = {name: name for name in parts}
    def assign(part, label):
        assert part in parts
        assert labels[part] in [part, label], 'Conflicting native ownership alias'
        labels[part] = label
    presentation = n.root(parts, 'ppt/presentation.xml')
    refs = presentation.findall('p:sldIdLst/p:sldId', NS)
    assert len(refs) == 3 and [x.get('id') for x in refs] == ['256', '257', '258']
    edges = n.relations(parts, 'ppt/presentation.xml')
    assert presentation.find('p:sldSz', NS).get('cx') == '9144000'
    assert presentation.find('p:sldSz', NS).get('cy') == '6858000'
    assign(n.target(parts, 'ppt/presentation.xml', 'theme'), 'theme/presentation')
    masters = presentation.findall('p:sldMasterIdLst/p:sldMasterId', NS)
    assert len(masters) == 3
    owners = []
    for ref in masters:
        master = edges[ref.get('{'+R+'}id')]['target']
        name = n.root(parts, master).find('p:cSld', NS).get('name')
        owners.append(name)
        assign(n.target(parts, master, 'theme'), 'theme/presentation' if name == 'Master' else 'theme/master/'+name)
    assert owners == ['Master', 'red', 'blue'], owners
    notes_master = n.target(parts, 'ppt/presentation.xml', 'notesMaster')
    assign(n.target(parts, notes_master, 'theme'), 'theme/notes')
    charts, images, workbooks = set(), set(), set()
    for i, ref in enumerate(refs):
        slide = edges[ref.get('{'+R+'}id')]['target']
        assert slide == f'ppt/slides/slide{i+1}.xml'
        root = n.root(parts, slide)
        assert root.get('show', '1') == ('0' if i == 1 else '1')
        assert root.get('showMasterSp', '1') == ('0' if i == 1 else '1')
        layout = n.target(parts, slide, 'slideLayout')
        master = n.target(parts, layout, 'slideMaster')
        assert n.root(parts, layout).find('p:cSld', NS).get('name') == ('Source' if i == 2 else 'Target')
        assert n.root(parts, master).find('p:cSld', NS).get('name') == ('red' if i == 2 else 'blue')
        assert [ph.attrib for ph in root.findall('.//p:ph', NS)] == [
            {'type': 'title', 'idx': '1' if i == 2 else '3'}, {'type': 'body', 'idx': '2'}, {'type': 'pic', 'idx': '3' if i == 2 else '1'}]
        frames = {'title': (40, 40, 450, 50), 'body': (40, 120, 320, 90), 'image': (450, 200, 192, 96)} if i == 1 else {'title': (20, 20, 500, 50), 'body': (20, 90, 350, 100), 'image': (430, 100, 128, 64)}
        objects = {node.find('.//p:cNvPr', NS).get('name'): node for node in root.find('p:cSld/p:spTree', NS) if node.find('.//p:cNvPr', NS) is not None}
        for name, frame in frames.items():
            node = objects[name+'-'+str(i)]; transform = node.find('p:spPr/a:xfrm', NS)
            off, ext = transform.find('a:off', NS), transform.find('a:ext', NS)
            assert tuple(int(value) for value in [off.get('x'), off.get('y'), ext.get('cx'), ext.get('cy')]) == tuple(value*12700 for value in frame), 'Wrong selected native placeholder frame'
        title_run = objects['title-'+str(i)].find('.//a:rPr', NS)
        body_run = objects['body-'+str(i)].find('.//a:rPr', NS)
        assert title_run.get('sz') == ('2800' if i == 1 else '2400')
        assert body_run.get('sz') == ('2000' if i == 1 else '1800')
        assert body_run.get('b') == ('1' if i == 1 else '0'), 'Text reset failed to resume target bold'
        assert body_run.find('a:hlinkClick', NS) is not None, 'Text reset lost its action'
        assert len(root.findall('.//a:tbl', NS)) == 1
        assert len(root.findall('.//{'+h.C+'}chart')) == 1
        assert not any(node.get('name') in ['red-anchor', 'blue-anchor'] for node in root.findall('.//p:cNvPr', NS)), 'Master artwork flattened onto a page'
        note = n.target(parts, slide, 'notesSlide')
        assert n.target(parts, note, 'slide') == slide
        assert n.target(parts, note, 'notesMaster') == notes_master
        assert ''.join(n.root(parts, note).itertext()) == 'Notes '+str(i)
        image = n.target(parts, slide, 'image'); images.add(image); assign(image, 'asset/logo')
        chart = n.target(parts, slide, 'chart'); charts.add(chart)
        workbook = n.target(parts, chart, 'package'); workbooks.add(workbook); assign(workbook, 'workbook/'+str(i))
    assert len(charts) == len(workbooks) == 3 and len(images) == 1
    assert {p for p in parts if p.startswith('ppt/charts/') and p.endswith('.xml')} == charts
    assert {p for p in parts if p.startswith('ppt/embeddings/')} == workbooks
    assert {p for p in parts if p.startswith('ppt/media/')} == images
    assert len(set(labels.values())) == len(labels), 'Aliased package identities'
    return labels


def exact_xml(node, xml):
    assert node is not None, 'Missing expected native addition'
    assert n.plain(node) == n.plain(E.fromstring(xml)), 'Unexpected native addition '+node.tag


def canonical(parts, native):
    labels = identity(parts)
    result, guids = {}, set()
    def edge(part, rid):
        value = n.relations(parts, part)[rid]
        target = value['target']
        if value['mode'] == 'Internal':
            assert target in parts, 'Dangling internal relationship'
            target = labels[target]
        return value['type'], value['mode'], target
    def tree(node, part, parent=None):
        attrs = dict(node.attrib)
        for key, value in list(attrs.items()):
            if key.startswith('{'+R+'}'):
                attrs[key] = repr(edge(part, value))
        if node.tag == '{'+A+'}pPr' and attrs.get('indent') == '-0': attrs['indent'] = '0'
        if node.tag == '{'+A+'}tcPr' and attrs.get('anchor') == 't': del attrs['anchor']
        if node.tag == '{'+P+'}sldSz' and native:
            assert attrs.pop('type') == 'screen4x3'
        children = []
        for child in node:
            if native and child.tag in ['{'+A+'}spLocks', '{'+A+'}picLocks']:
                ph = parent.find('p:nvPr/p:ph', NS) if parent is not None else None
                assert ph is not None and node.tag in ['{'+P+'}cNvSpPr', '{'+P+'}cNvPicPr']
                expected = {'noGrp': '1', 'noRot': '1', 'noChangeAspect': '1'} if ph.get('type') == 'sldImg' else {'noGrp': '1'}
                assert child.attrib == expected and not len(child)
                continue
            if native and child.tag == '{'+P+'}extLst' and part == 'ppt/presentation.xml':
                exact_xml(child, '<p:extLst xmlns:p="'+P+'" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst/></p:ext></p:extLst>')
                continue
            children.append(tree(child, part, node))
        return node.tag, sorted(attrs.items()), node.text or '', tuple(children)
    for part, body in parts.items():
        if part.endswith('.rels'):
            owner = '' if part == '_rels/.rels' else str(Path(part).parent.parent/Path(part).name.removesuffix('.rels'))
            relations = n.relations(parts, owner)
            xml = E.fromstring(body)
            assert xml.tag == '{'+REL+'}Relationships' and not xml.attrib
            assert len(relations) == len(xml) and all(e.tag == '{'+REL+'}Relationship' and {'Id', 'Type', 'Target'} <= set(e.attrib) <= {'Id', 'Type', 'Target', 'TargetMode'} for e in xml)
            result[labels[part]] = sorted(edge(owner, rid) for rid in relations)
        elif part.startswith('ppt/charts/') and part.endswith('.xml'):
            result[labels[part]] = h.canonical(body, native, guids)
        elif not part.endswith('.xml'):
            result[labels[part]] = a.sha(body)
        else:
            root = E.fromstring(body)
            if part == '[Content_Types].xml':
                for node in root:
                    if 'PartName' in node.attrib: node.set('PartName', '/'+labels[node.get('PartName')[1:]])
                root[:] = sorted(root, key=lambda node: (node.tag, sorted(node.attrib.items())))
            if native:
                if part == 'ppt/viewProps.xml':
                    view = root.find('p:slideViewPr/p:cSldViewPr/p:cViewPr', NS)
                    for axis in ['sx', 'sy']:
                        node = view.find('p:scale/a:'+axis, NS)
                        assert node.attrib == {'n': '128', 'd': '100'}
                        node.set('n', '136')
                    node = view.find('p:origin', NS)
                    assert node.attrib == {'x': '1704', 'y': '176'}
                    node.set('x', '216'); node.set('y', '312')
                for table in root.iter('{'+A+'}tbl'):
                    for tag, nodes, start, uri, local in [
                        ('gridCol', table.findall('a:tblGrid/a:gridCol', NS), 20000, '{9D8B030D-6E8A-4147-A177-3AD203B41FA5}', 'colId'),
                        ('tr', table.findall('a:tr', NS), 10000, '{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}', 'rowId')]:
                        assert len(nodes) == 2
                        for index, node in enumerate(nodes):
                            extension = node.find('a:extLst', NS)
                            assert extension is not None
                            exact_xml(extension, '<a:extLst xmlns:a="'+A+'" xmlns:a16="'+A16+'"><a:ext uri="'+uri+'"><a16:'+local+' val="'+str(start+index)+'"/></a:ext></a:extLst>')
                            node.remove(extension)
                if part == 'ppt/presProps.xml':
                    exact_xml(root, '<p:presentationPr xmlns:p="'+P+'" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"><p:extLst><p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData val="0"/></p:ext><p:ext uri="{D31A062A-798A-4329-ABDD-BBA856620510}"><p14:defaultImageDpi val="220"/></p:ext><p:ext uri="{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}"><p15:chartTrackingRefBased val="0"/></p:ext></p:extLst></p:presentationPr>')
                    root.clear()
                if part == 'docProps/core.xml':
                    for tag in ['{'+DCT+'}modified', '{'+CP+'}lastPrinted', '{'+CP+'}lastModifiedBy']:
                        assert len(root.findall(tag)) == 1, 'Unexpected metadata field cardinality'
                    for child in list(root):
                        if child.tag in ['{'+DCT+'}modified', '{'+CP+'}lastPrinted']:
                            assert re.fullmatch(r'\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ', child.text)
                            datetime.fromisoformat(child.text.replace('Z', '+00:00'))
                            assert child.attrib == ({'{http://www.w3.org/2001/XMLSchema-instance}type': 'dcterms:W3CDTF'} if child.tag == '{'+DCT+'}modified' else {})
                            root.remove(child)
                        elif child.tag == '{'+CP+'}lastModifiedBy':
                            assert child.text == 'openHacking' and not child.attrib and not len(child); root.remove(child)
                if part == 'docProps/app.xml': root = app_properties(root)
            result[labels[part]] = tree(root, part)
    assert not native or len(guids) == 3
    return result


def app_properties(root):
    simple = {'TotalTime': '0', 'Words': '39', 'Application': 'Microsoft Macintosh PowerPoint', 'PresentationFormat': '全屏显示(4:3)', 'Paragraphs': '24', 'Slides': '3', 'Notes': '3', 'HiddenSlides': '1', 'MMClips': '0', 'ScaleCrop': 'false', 'LinksUpToDate': 'false', 'SharedDoc': 'false', 'HyperlinksChanged': 'false', 'AppVersion': '16.0000'}
    assert not root.attrib and len(root) == len(simple)+2
    for key, value in simple.items():
        nodes = root.findall('{'+EP+'}'+key); assert len(nodes) == 1
        node = nodes[0]
        assert (re.fullmatch(r'\d+', node.text) if key == 'TotalTime' else node.text == value) and not node.attrib and not len(node), key
    expected = '<HeadingPairs xmlns="'+EP+'" xmlns:vt="'+VT+'"><vt:vector size="6" baseType="variant">'+''.join('<vt:variant><vt:lpstr>'+name+'</vt:lpstr></vt:variant><vt:variant><vt:i4>'+str(count)+'</vt:i4></vt:variant>' for name, count in [('已用的字体', 2), ('主题', 3), ('幻灯片标题', 3)])+'</vt:vector></HeadingPairs>'
    exact_xml(root.find('{'+EP+'}HeadingPairs'), expected)
    expected = '<TitlesOfParts xmlns="'+EP+'" xmlns:vt="'+VT+'"><vt:vector size="8" baseType="lpstr">'+''.join('<vt:lpstr>'+name+'</vt:lpstr>' for name in ['Arial', 'Courier New', 'PPTKit', 'PPTKit', 'PPTKit', 'TITLE 0', 'TITLE 1', 'TITLE 2'])+'</vt:vector></TitlesOfParts>'
    exact_xml(root.find('{'+EP+'}TitlesOfParts'), expected)
    return E.fromstring('<Properties xmlns="'+EP+'"><Application>PPTKit</Application><Slides>3</Slides><PresentationFormat>On-screen Show (4:3)</PresentationFormat></Properties>')


def review(parts, source):
    expected, actual = canonical(source, False), canonical(parts, True)
    assert expected.keys() == actual.keys(), 'Changed package part identity set'
    changed = [key for key in expected if expected[key] != actual[key]]
    assert not changed, ('Unapproved first-save change', changed)
    return a.sha(json.dumps(actual, sort_keys=True).encode())


def reset_pdf(file, before_file):
    document, original = fitz.open(file), fitz.open(before_file)
    assert len(document) == len(original) == 3
    page, before = document[1], original[1]
    assert list(page.rect) == [0.0, 0.0, 842.0, 595.0]
    assert page.get_text() == before.get_text(), 'Reset lost or duplicated actual text'
    spans = page.get_texttrace()
    def text(span): return ''.join(chr(char[0]) for char in span['chars'])
    assert all(span['font'] in ['CourierNewPSMT', 'CourierNewPS-BoldMT'] for span in spans), 'Target theme font not used'
    for value, face, size, origin in [
        ('TITLE 1', 'CourierNewPSMT', 28.999990463256836, (96.63194274902344, 88.7999267578125)),
        ('BODY 1', 'CourierNewPS-BoldMT', 21.0, (91.45832824707031, 163.199951171875)),
        ('Ordinary 1', 'CourierNewPSMT', 18.999996185302734, (75.9375, 297.35992431640625)),
    ]:
        matches = [span for span in spans if text(span) == value]; assert len(matches) == 1
        span = matches[0]
        assert span['font'] == face and span['size'] == size and span['chars'][0][2] == origin, 'Wrong actual target text/font/placement'
    # Printed font sizes above are observed scaled PDF values. Source/native XML
    # independently retains 28/20/18 pt; do not mistake print scaling for reset.
    drawings, old = page.get_drawings(), before.get_drawings()
    assert len(drawings) == len(old) == 18, 'Hidden master artwork unexpectedly appears'
    assert tuple(drawings[1]['rect']) == (48.0, 18.0, 793.0, 577.0)
    assert drawings[1]['fill'] == (0.9411764740943909, 0.9725490212440491, 1.0), 'Target background not resumed'
    for index in [3, 4, 5, 6]:
        x, y = copy.deepcopy(drawings[index]), copy.deepcopy(old[index])
        assert x.pop('fill') == (0.0, 0.2666666805744171, 0.800000011920929)
        assert y.pop('fill') == (0.800000011920929, 0.13331808149814606, 0.0)
        assert x == y, 'Ordinary/table geometry or opacity reset unexpectedly'
    assert tuple(drawings[-1]['rect']) == (420.5, 188.80560302734375, 513.625, 274.7259521484375), 'Connector did not follow new placeholder frames'
    images = page.get_image_info(hashes=True)
    assert len(images) == 1 and images[0]['width'] == 8 and images[0]['height'] == 4
    assert images[0]['digest'] == before.get_image_info(hashes=True)[0]['digest'], 'Reset changed image content'
    assert images[0]['bbox'] == (513.625, 225.03701782226562, 712.2916870117188, 324.414794921875), 'Picture did not move/resize to target frame'
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    assert pix.pixel(100, 40) == (240, 248, 255)
    assert pix.pixel(1440, 120) == (240, 248, 255), 'Hidden target master art rendered'
    rect = fitz.Rect(images[0]['bbox'])
    for u, v, color in [(0.2, 0.2, (36, 87, 214)), (0.2, 0.8, (34, 197, 94)), (0.8, 0.2, (217, 74, 102)), (0.8, 0.8, (250, 204, 21))]:
        assert pix.pixel(round((rect.x0+rect.width*u)*2), round((rect.y0+rect.height*v)*2)) == color, 'Picture occluded or wrongly cropped'
    body = next(span for span in spans if text(span) == 'BODY 1')
    rect = fitz.Rect(body['bbox'])
    assert any((color := pix.pixel(x, y))[2] > color[0]+50 and color[2] > color[1]+20 for x in range(int(rect.x0*2), int(rect.x1*2)) for y in range(int(rect.y0*2), int(rect.y1*2))), 'Body text occluded'
    return a.pdf_page_record(document, 1)


def main():
    manifest = json.loads(Path('docs/testing/layout-rebind-evidence.json').read_text())
    directory = Path(manifest['nativeDirectory'])
    source = n.payload(directory/'input.pptx')
    assert a.sha((directory/'input.pptx').read_bytes()) == manifest['sourceSha256']
    rounds = []
    for i in range(1, 4):
        file = directory/f'rebind-round{i}.pptx'
        rounds.append({'round': i, 'file': str(file), 'sha256': a.sha(file.read_bytes()), 'canonicalSha256': review(n.payload(file), source)})
    reset = reset_pdf(directory/'first.pdf', directory/'before.pdf')
    for i in range(1, 4): assert reset_pdf(directory/f'rebind-round{i}.pdf', directory/'before.pdf') == reset
    positive = n.payload(directory/'rebind-round1.pptx')
    out = Path(tempfile.mkdtemp(prefix='first-save-negatives-', dir=directory))
    negatives = []
    def xml_patch(part, callback):
        def mutate(parts):
            root = E.fromstring(parts[part]); callback(root); parts[part] = E.tostring(root)
        return mutate
    def relation(root, kind): return next(node for node in root if node.get('Type').endswith('/'+kind))
    def changed_attr(path, attr, value):
        def mutate(root):
            node = root.find(path, NS); assert node is not None; node.set(attr, value)
        return mutate
    def body_shape(root):
        return next(node for node in root.findall('.//p:sp', NS) if node.find('p:nvSpPr/p:cNvPr', NS).get('name') == 'body-1')
    cases = [
        ('wrong-layout', xml_patch('ppt/slides/_rels/slide3.xml.rels', lambda root: relation(root, 'slideLayout').set('Target', '../slideLayouts/slideLayout3.xml'))),
        ('wrong-theme-owner', xml_patch('ppt/slideMasters/_rels/slideMaster2.xml.rels', lambda root: relation(root, 'theme').set('Target', '../theme/theme3.xml'))),
        ('wrong-placeholder-index', xml_patch('ppt/slides/slide1.xml', changed_attr('.//p:ph', 'idx', '1'))),
        ('unhide-page', xml_patch('ppt/slides/slide2.xml', lambda root: root.set('show', '1'))),
        ('show-inherited-artwork', xml_patch('ppt/slides/slide2.xml', lambda root: root.set('showMasterSp', '1'))),
        ('wrong-notes-backlink', xml_patch('ppt/notesSlides/_rels/notesSlide1.xml.rels', lambda root: relation(root, 'slide').set('Target', '../slides/slide2.xml'))),
        ('dangling-image', xml_patch('ppt/slides/_rels/slide1.xml.rels', lambda root: relation(root, 'image').set('Target', '../media/missing.png'))),
        ('unknown-relationship-attribute', xml_patch('_rels/.rels', lambda root: root[0].set('unknown', '1'))),
        ('unknown-relationship-root-attribute', xml_patch('_rels/.rels', lambda root: root.set('unknown', '1'))),
        ('duplicate-relationship-id', xml_patch('ppt/slides/_rels/slide1.xml.rels', lambda root: root[0].set('Id', root[1].get('Id')))),
        ('title-font', xml_patch('ppt/slides/slide1.xml', changed_attr('.//p:sp/p:txBody/a:p/a:r/a:rPr/a:latin', 'typeface', 'Wrong Face'))),
        ('reset-body-bold', xml_patch('ppt/slides/slide2.xml', lambda root: body_shape(root).find('.//a:rPr', NS).set('b', '0'))),
        ('reset-body-action', xml_patch('ppt/slides/slide2.xml', lambda root: body_shape(root).find('.//a:rPr', NS).remove(body_shape(root).find('.//a:hlinkClick', NS)))),
        ('paragraph-indent', xml_patch('ppt/slides/slide1.xml', changed_attr('.//a:pPr', 'indent', '1'))),
        ('table-paragraph-font', xml_patch('ppt/slides/slide1.xml', changed_attr('.//a:tbl//a:endParaRPr/a:latin', 'typeface', 'Courier New'))),
        ('table-row-id', xml_patch('ppt/slides/slide1.xml', lambda root: next(root.iter('{'+A16+'}rowId')).set('val', '10099'))),
        ('table-border', xml_patch('ppt/slides/slide1.xml', lambda root: root.find('.//a:tcPr', NS).remove(root.find('.//a:tcPr/a:lnL', NS)))),
        ('placeholder-lock', xml_patch('ppt/slides/slide1.xml', changed_attr('.//a:spLocks', 'noGrp', '0'))),
        ('one-emu-geometry', xml_patch('ppt/slides/slide1.xml', changed_attr('.//p:sp/p:spPr/a:xfrm/a:ext', 'cx', '6350001'))),
        ('chart-cache', xml_patch('ppt/charts/chart1.xml', lambda root: root.find('.//c:numCache/c:pt/c:v', h.NS).__setattr__('text', '99'))),
        ('chart-guid', xml_patch('ppt/charts/chart1.xml', lambda root: next(root.iter('{'+h.C16+'}uniqueId')).set('val', 'bad'))),
        ('unknown-presentation-extension', xml_patch('ppt/presentation.xml', lambda root: E.SubElement(root.find('p:extLst', NS), '{'+P+'}ext', {'uri': 'unknown'}))),
        ('presentation-aspect', xml_patch('ppt/presentation.xml', changed_attr('p:sldSz', 'type', 'screen16x9'))),
        ('default-text-size', xml_patch('ppt/presentation.xml', changed_attr('p:defaultTextStyle/a:lvl1pPr/a:defRPr', 'sz', '1900'))),
        ('app-hidden-count', xml_patch('docProps/app.xml', lambda root: root.find('{'+EP+'}HiddenSlides').__setattr__('text', '0'))),
        ('unknown-clock-namespace', xml_patch('docProps/core.xml', lambda root: E.SubElement(root, '{urn:unknown}lastPrinted').__setattr__('text', '2026-10-27T00:00:00Z'))),
        ('duplicate-clock', xml_patch('docProps/core.xml', lambda root: root.append(copy.deepcopy(root.find('{'+CP+'}lastPrinted'))))),
    ]
    def workbook(parts):
        part = n.target(parts, 'ppt/charts/chart1.xml', 'package'); parts[part] += b'changed'
    def image(parts):
        part = n.target(parts, 'ppt/slides/slide1.xml', 'image'); parts[part] += b'changed'
    cases += [('workbook-payload', workbook), ('image-payload', image)]
    for name, mutate in cases:
        parts = copy.copy(positive); mutate(parts); file = out/(name+'.pptx')
        with zipfile.ZipFile(file, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            for part, body in parts.items(): archive.writestr(part, body)
        try: review(n.payload(file), source)
        except (AssertionError, KeyError) as error:
            negatives.append({'name': name, 'file': str(file), 'bytes': file.stat().st_size, 'sha256': a.sha(file.read_bytes()), 'reason': str(error)})
        else: raise AssertionError('Accepted corrupt actual package '+name)
    pdf_negatives = []
    for name in ['unchanged-reset', 'covered-reset-picture', 'covered-reset-body']:
        doc = fitz.open(directory/('before.pdf' if name == 'unchanged-reset' else 'first.pdf'))
        if name != 'unchanged-reset':
            rect = fitz.Rect(doc[1].get_image_info()[0]['bbox']) if name == 'covered-reset-picture' else doc[1].search_for('BODY 1')[0]
            doc[1].draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))
        file = out/(name+'.pdf'); doc.save(file)
        try: reset_pdf(file, directory/'before.pdf')
        except AssertionError as error: pdf_negatives.append({'name': name, 'file': str(file), 'bytes': file.stat().st_size, 'sha256': a.sha(file.read_bytes()), 'reason': str(error)})
        else: raise AssertionError('Accepted corrupt actual reset PDF '+name)
    dependencies = [Path(__file__), HERE/'master-chart-first-save.py', HERE/'master-cell-properties-native.py', HERE/'master-artwork-native.py', HERE/'multi-master-native.py', HERE/'multi-master-operations.py']
    result = {'schemaVersion': 1, 'recordedAt': datetime.now(timezone.utc).isoformat(), 'verified': True, 'verifiedNamedResetPage': True, 'directory': str(directory), 'sourceSha256': a.sha((directory/'input.pptx').read_bytes()), 'rounds': rounds, 'resetPageRecord': reset, 'negatives': negatives+pdf_negatives, 'dependencies': [{'file': str(file), 'sha256': a.sha(file.read_bytes())} for file in dependencies], 'scope': 'All package parts for the named three-page fresh-output rebind control compared to its exact source. Relationship IDs/order, owner-bound theme/image/workbook part renames, validated placeholder locks, table row/column IDs, zero indentation/default top anchor, explicit native 4:3 type/empty guides, named view/presProps and validated metadata changes, and captured positive line-chart first-save defaults are the only admitted differences. All other XML/binary values and ownership/resource aliases compare exactly. Named hidden reset-page physical checks cover unchanged strings/actions, actual target fonts/bold/positions, background and hidden master artwork, local geometry/opacity, picture content/frame/visibility and connector movement. No general-purpose OOXML normalization or source editing authority.'}
    Path('docs/testing/layout-rebind-first-save-evidence.json').write_text(json.dumps(result, indent=2)+'\n')
    print(json.dumps({'verified': True, 'rounds': len(rounds), 'actualZipNegatives': len(negatives), 'actualPdfNegatives': len(pdf_negatives)}))


if __name__ == '__main__': main()
