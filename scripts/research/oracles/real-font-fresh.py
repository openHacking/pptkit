"""Independent named fresh-author PPTX and native PDF evidence, not a general importer."""
import hashlib, json, posixpath, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image, ImageChops, ImageDraw

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def sha(file):
    return hashlib.sha256(Path(file).read_bytes()).hexdigest()
def inspect_package(file, expected, source=False):
    if source:
        assert sha(file) == expected['sourceSha256'], 'source bytes changed'
    rows = []
    with zipfile.ZipFile(file) as z:
        assert len(z.namelist()) == len(set(z.namelist())), 'duplicate ZIP entry'
        rels = {r.attrib['Id']: r for r in ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))}
        presentation = ET.fromstring(z.read('ppt/presentation.xml'))
        slides = presentation.findall('p:sldIdLst/p:sldId', NS)
        assert len(slides) == len(expected['controls']) == 16
        for slide, control in zip(slides, expected['controls']):
            rel = rels[slide.attrib['{' + NS['r'] + '}id']]
            assert rel.attrib['Type'].endswith('/slide') and rel.attrib.get('TargetMode') != 'External'
            target = rel.attrib['Target']
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('ppt', target))
            tree = ET.fromstring(z.read(part))
            shapes = {}
            for shape in tree.findall('p:cSld/p:spTree/p:sp', NS):
                name = shape.find('p:nvSpPr/p:cNvPr', NS).attrib['name']
                assert name not in shapes
                shapes[name] = shape
            assert set(shapes) == {'Label', 'Probe'}
            assert ''.join(shapes['Label'].itertext()) == control['id'], 'slide order/label changed'
            body = shapes['Probe'].find('p:txBody', NS)
            assert body.find('a:bodyPr', NS).attrib['wrap'] == 'none'
            runs = body.findall('a:p/a:r', NS)
            assert len(runs) == 1, 'unedited control topology changed'
            run = runs[0]; style = run.find('a:rPr', NS)
            assert run.find('a:t', NS).text == control['text'], 'body text changed'
            assert style.attrib['lang'] == control['language']
            assert style.attrib['sz'] == '4800'
            slots = [style.find('a:' + slot, NS).attrib['typeface'] for slot in ['latin', 'ea', 'cs']]
            assert slots == control['expectedSlots'], 'font slots changed'
            rows.append({'id': control['id'], 'part': part, 'text': control['text'], 'language': style.attrib['lang'], 'slots': slots, 'runs': len(runs)})
    return {'sha256': sha(file), 'controls': rows}

def inspect_pdf(file, expected, output=None, baseline=None):
    doc = fitz.open(file)
    assert len(doc) == len(expected['controls']) == 16, 'PDF page count'
    fonts = {f['id']: f['postscriptName'] for f in expected['fonts']}
    rows, images = [], []
    previous = fitz.open(baseline) if baseline else None
    if previous: assert len(previous) == len(doc)
    for index, (page, control) in enumerate(zip(doc, expected['controls'])):
        assert abs(page.rect.width - 720) < .1 and abs(page.rect.height - 260) < .1
        label = ''.join(s['text'] for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['bbox'][1] < 60).strip()
        assert label == control['id'], 'PDF order/label changed'
        traces = [t for t in page.get_texttrace() if t['bbox'][1] > 60]
        actual = sorted(set(t['font'] for t in traces))
        assert actual == [fonts[control['fontId']]], 'actual PDF face differs: ' + control['id']
        glyphs = [c for t in traces for c in t['chars']]
        assert glyphs and all(c[1] != 0 for c in glyphs), 'missing glyph zero'
        pix = page.get_pixmap(alpha=False)
        image = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
        body = image.crop((0, 60, 720, 260))
        ink = sum(min(pixel) < 200 for pixel in body.getdata())
        assert ink > 100, 'body invisible'
        if previous:
            old = previous[index].get_pixmap(alpha=False)
            assert pix.width == old.width and pix.height == old.height and pix.samples == old.samples, 'native raster changed'
        rows.append({'id': control['id'], 'fonts': actual, 'glyphCount': len(glyphs), 'glyphZeroCount': 0, 'darkPixels': ink, 'rasterSha256': hashlib.sha256(pix.samples).hexdigest()})
        images.append(image)
    if output:
        sheet = Image.new('RGB', (720*2, 260*8), 'white')
        for i, image in enumerate(images): sheet.paste(image, ((i%2)*720, (i//2)*260))
        sheet.save(output)
    return {'sha256': sha(file), 'controls': rows, 'baselineRasterMatch': bool(previous)}

if __name__ == '__main__':
    directory = Path(sys.argv[1]); expected = json.loads((directory/'expected.json').read_text())
    for font in expected['fonts']: assert sha(font['file']) == font['sha256'], 'installed font changed'
    report = {'source': inspect_package(directory/'fresh-fonts.pptx', expected, True), 'pdf': inspect_pdf(directory/'fresh-fonts.pdf', expected, directory/'contact-sheet.png')}
    for name in sys.argv[2:]:
        report[name] = {'package': inspect_package(directory/(name+'.pptx'), expected), 'pdf': inspect_pdf(directory/(name+'.pdf'), expected, baseline=directory/'fresh-fonts.pdf')}
    (directory/'inspection.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n')
    print(json.dumps({'controls': 16, 'nativeFacesMatch': 16, 'rounds': sys.argv[2:], 'directory': str(directory)}))
