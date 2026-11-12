"""Strict, editing-specific oracle; the unedited table oracle stays unchanged."""
import datetime, hashlib, importlib.util, json, math, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image

spec = importlib.util.spec_from_file_location('table_baseline', Path(__file__).with_name('native-table-fonts.py'))
baseline = importlib.util.module_from_spec(spec)
spec.loader.exec_module(baseline)
NS = baseline.NS
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
MOD_URI = '{D42A27DB-BD31-4B8C-83A1-F6EECF244321}'


def sha(file):
    return hashlib.sha256(Path(file).read_bytes()).hexdigest()


def tree(node):
    return (node.tag, sorted(node.attrib.items()), node.text, node.tail, [tree(child) for child in node])


def body_traces(page):
    return [trace for trace in page.get_texttrace() if trace['bbox'][1] > 60]


def inspect_package(file, directory, expected):
    original = directory / 'baseline.pptx'
    assert sha(original) == expected['nativeBaseline']['sha256'], 'baseline hash'
    controls = baseline.inspect_package(original, expected)['controls']
    targets = expected['editTargets']
    assert targets == ([3, 5, 6, 7, 12, 14, 15, 21, 25, 26] if len(controls) == 27 else [3, 5, 6, 7, 8])
    changed = {controls[i - 1]['part']: i for i in targets}
    rows, metadata = [], []
    with zipfile.ZipFile(original) as source, zipfile.ZipFile(file) as candidate:
        names = candidate.namelist()
        assert len(names) == len(set(names)), 'duplicate ZIP entry'
        assert set(names) == set(source.namelist()), 'part inventory changed'
        for part in names:
            before, after = source.read(part), candidate.read(part)
            if part in ['docProps/app.xml', 'docProps/core.xml']:
                left, right = ET.fromstring(before), ET.fromstring(after)
                allowed = {'TotalTime', 'Words'} if part.endswith('app.xml') else {'revision', 'modified'}
                assert [c.tag for c in left] == [c.tag for c in right], 'metadata inventory'
                for old, new in zip(left, right):
                    key = old.tag.split('}')[-1]
                    if key in allowed:
                        assert old.attrib == new.attrib and len(old) == len(new) == 0
                        if key == 'Words': assert int(new.text) == int(old.text) + len(targets), 'word count'
                        elif key == 'revision': assert int(new.text) == int(old.text) + 1, 'revision'
                        elif key == 'TotalTime': assert int(new.text) >= int(old.text), 'editing time'
                        else: assert datetime.datetime.fromisoformat(new.text.replace('Z', '+00:00')) >= datetime.datetime.fromisoformat(old.text.replace('Z', '+00:00')), 'modified time'
                        metadata.append({'part': part, 'property': key, 'before': old.text, 'after': new.text})
                        old.text = new.text
                assert tree(left) == tree(right), 'unexpected metadata change'
                continue
            if part not in changed:
                assert before == after, 'unrelated part changed: ' + part
                continue
            index = changed[part]
            control = expected['controls'][index - 1]
            left, right = ET.fromstring(before), ET.fromstring(after)
            frames = right.findall('p:cSld/p:spTree/p:graphicFrame', NS)
            assert len(frames) == 1
            frame = frames[0]
            target_frame = left.find('p:cSld/p:spTree/p:graphicFrame', NS)
            para = target_frame.find('.//a:tbl/a:tr/a:tc/a:txBody/a:p', NS)
            runs = para.findall('a:r', NS)
            assert len(runs) == 1
            runs[0].find('a:rPr', NS).set('dirty', '0')
            added = ET.Element('{' + NS['a'] + '}r')
            attrs = {'lang': 'el-GR', 'sz': '4800', 'dirty': '0'}
            if control['sampleId'] not in ['ar', 'he']: attrs['altLang'] = control['language']
            ET.SubElement(added, '{' + NS['a'] + '}rPr', attrs)
            ET.SubElement(added, '{' + NS['a'] + '}t').text = 'Ω'
            para.insert(list(para).index(runs[0]) + 1, added)
            para.find('a:endParaRPr', NS).set('dirty', '0')
            nvpr = frame.find('p:nvGraphicFramePr/p:nvPr', NS)
            expected_nvpr = target_frame.find('p:nvGraphicFramePr/p:nvPr', NS)
            assert len(expected_nvpr) == 0
            mod = nvpr.find('p:extLst/p:ext/{' + P14 + '}modId', NS)
            assert mod is not None and set(mod.attrib) == {'val'} and len(mod) == 0
            assert mod.attrib['val'].isdigit() and 0 <= int(mod.attrib['val']) <= 2**32 - 1
            extlist = ET.SubElement(expected_nvpr, '{' + NS['p'] + '}extLst')
            ext = ET.SubElement(extlist, '{' + NS['p'] + '}ext', {'uri': MOD_URI})
            ET.SubElement(ext, '{' + P14 + '}modId', dict(mod.attrib))
            assert tree(left) == tree(right), 'unexpected edited slide tree: ' + part
            actual_para = frame.find('.//a:tbl/a:tr/a:tc/a:txBody/a:p', NS)
            actual_runs = actual_para.findall('a:r', NS)
            rows.append({'id': control['id'], 'part': part, 'slide': index, 'modId': mod.attrib['val'], 'runs': [{'text': r.find('a:t', NS).text, 'attributes': dict(r.find('a:rPr', NS).attrib)} for r in actual_runs], 'localFontOverrides': [], 'declared': controls[index - 1]['declared'], 'selector': controls[index - 1]['selector']})
    return {'sha256': sha(file), 'controls': sorted(rows, key=lambda row: row['slide']), 'metadataChanges': metadata, 'unrelatedPartsByteExact': True}


def inspect_pdf(file, directory, expected, previous=None):
    assert sha(directory / 'baseline.pdf') == expected['nativeBaseline']['pdfSha256'], 'baseline PDF hash'
    rows = []
    with fitz.open(file) as doc, fitz.open(directory / 'baseline.pdf') as source:
        assert len(doc) == len(source) == len(expected['controls'])
        prior = fitz.open(previous) if previous else None
        try:
            for index, (page, old, control) in enumerate(zip(doc, source, expected['controls']), 1):
                assert page.rect == old.rect and abs(page.rect.width - 720) < .1 and abs(page.rect.height - 260) < .1
                label = ''.join(s['text'] for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['bbox'][1] < 60).strip()
                assert label == control['id'], 'PDF slide order'
                baseline.mixed.font_resources(doc, page)
                traces, before = body_traces(page), body_traces(old)
                glyphs = [(t['font'], c) for t in traces for c in t['chars']]
                assert glyphs and all(c[1] != 0 for _, c in glyphs), 'glyph zero'
                edited = index in expected['editTargets']
                if edited:
                    omega = [(f, c) for f, c in glyphs if c[0] == 937]
                    assert len(omega) == 1 and omega[0][0] == 'Georgia', 'appended actual font/glyph'
                    assert [(f, c) for f, c in glyphs if c[0] != 937] == [(t['font'], c) for t in before for c in t['chars']], 'original glyph/font/origin changed'
                else:
                    assert traces == before, 'unmodified text trace changed'
                pix = page.get_pixmap(alpha=False)
                if not edited:
                    assert pix.samples == old.get_pixmap(alpha=False).samples, 'unmodified page raster'
                if prior:
                    assert pix.samples == prior[index - 1].get_pixmap(alpha=False).samples, 'round raster changed'
                image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
                visible = []
                for trace in traces:
                    text = ''.join(chr(c[0]) for c in trace['chars'])
                    if not text.strip(): continue
                    b = trace['bbox']; bounds = (max(0, math.floor(b[0])), max(0, math.floor(b[1])), min(720, math.ceil(b[2])), min(260, math.ceil(b[3])))
                    pixels = image.crop(bounds).tobytes()
                    ink = sum(min(pixels[i:i+3]) < 200 for i in range(0, len(pixels), 3))
                    assert ink > 5, 'invisible trace'
                    visible.append({'font': trace['font'], 'text': text, 'darkPixels': ink, 'bounds': bounds})
                rows.append({'id': control['id'], 'edited': edited, 'actualFonts': sorted(set(t['font'] for t in traces)), 'sourceText': control['text'] + ('Ω' if edited else ''), 'pdfText': ''.join(chr(c[0]) for _, c in glyphs), 'originalGlyphsAndOriginsUnchanged': True, 'visibleTraceInk': visible, 'rasterSha256': hashlib.sha256(pix.samples).hexdigest()})
        finally:
            if prior: prior.close()
    return {'sha256': sha(file), 'controls': rows, 'nativeRoundRasterMatch': bool(previous)}


def inspect(directory):
    expected = json.loads((directory / 'expected.json').read_text())
    rounds = []
    for number in range(1, 4):
        stem = 'append-right-round' + str(number)
        package = inspect_package(directory / (stem + '.pptx'), directory, expected)
        if rounds: assert package['controls'] == rounds[0]['package']['controls'], 'run/selector/modId stability'
        pdf = inspect_pdf(directory / (stem + '.pdf'), directory, expected, directory / 'append-right-round1.pdf' if number > 1 else None)
        rounds.append({'name': stem, 'package': package, 'pdf': pdf})
    with fitz.open(directory / 'append-right-round1.pdf') as doc:
        sheet = Image.new('RGB', (720 * 3, 260 * math.ceil(len(doc) / 3)), 'white')
        for i, page in enumerate(doc):
            pix = page.get_pixmap(alpha=False)
            sheet.paste(Image.frombytes('RGB', (pix.width, pix.height), pix.samples), ((i % 3) * 720, (i // 3) * 260))
        sheet.save(directory / 'append-contact-sheet.png')
    result = {'schemaVersion': 1, 'directory': str(directory), 'baseline': expected['nativeBaseline'], 'editTargets': expected['editTargets'], 'rounds': rounds, 'scope': 'Named native append only; source declarations and original glyphs preserved. No bold/italic/theme/notes or complete bidi/mark-outline claim.'}
    (directory / 'append-inspection.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    return result


if __name__ == '__main__':
    result = inspect(Path(sys.argv[1]))
    print(json.dumps({'editedControls': len(result['editTargets']), 'pagesPerRound': len(result['rounds'][0]['pdf']['controls']), 'nativeRounds': 3}))
