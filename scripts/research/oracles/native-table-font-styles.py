"""Named native table style operations; original baseline and append oracles remain strict."""
import datetime, hashlib, importlib.util, json, math, re, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image

spec = importlib.util.spec_from_file_location('table_append', Path(__file__).with_name('native-table-font-edits.py'))
append = importlib.util.module_from_spec(spec)
spec.loader.exec_module(append)
NS = append.NS
STAGES = {'bold': {'b': '1'}, 'italic': {'b': '0', 'i': '1'}, 'regular': {'b': '0', 'i': '0'}}
BOLD = {'STSongti-TC-Regular': 'STSongti-TC-Bold', 'MalgunGothic': 'MalgunGothicBold', 'AppleSDGothicNeo-Regular': 'AppleSDGothicNeo-Bold', 'GeezaPro': 'GeezaPro-Bold', 'HiraginoSansGB-W3': 'HiraginoSansGB-W6', 'ArialMT': 'Arial-BoldMT', 'Batang': 'Batang', 'SimSun': 'SimSun', 'ArialUnicodeMS': 'ArialUnicodeMS'}
SYNTHETIC_BOLD = {'Batang', 'SimSun', 'ArialUnicodeMS'}
MATRIX = r'BT\s+(?:[-\d.]+\s+Tc\s+)?([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm\s+/([A-Za-z0-9]+)\s+1\s+Tf'


def verify_baseline(directory, expected):
    for extension in ['pptx', 'pdf']:
        assert append.sha(directory / ('baseline.' + extension)) == expected['styleBaseline'][extension], 'style baseline hash'
        original = Path(expected['appendSourceDirectory']) / ('append-right-round3.' + extension)
        assert append.sha(original) == expected['styleBaseline'][extension], 'original append evidence changed'
    return append.inspect_package(directory / 'baseline.pptx', Path(expected['appendSourceDirectory']), expected)


def inspect_package(file, directory, expected, stage):
    source_rows = verify_baseline(directory, expected)['controls']
    changed = {row['part']: row for row in source_rows}
    rows, metadata = [], []
    with zipfile.ZipFile(directory / 'baseline.pptx') as source, zipfile.ZipFile(file) as candidate:
        names = candidate.namelist()
        assert len(names) == len(set(names)) and set(names) == set(source.namelist()), 'part inventory'
        for part in names:
            before, after = source.read(part), candidate.read(part)
            if part in ['docProps/app.xml', 'docProps/core.xml']:
                left, right = ET.fromstring(before), ET.fromstring(after)
                assert [c.tag for c in left] == [c.tag for c in right]
                for old, new in zip(left, right):
                    key = old.tag.split('}')[-1]
                    if key in ({'TotalTime'} if part.endswith('app.xml') else {'revision', 'modified'}):
                        assert old.attrib == new.attrib and len(old) == len(new) == 0
                        if key == 'revision': assert int(new.text) == int(old.text) + list(STAGES).index(stage) + 1, 'stage revision'
                        elif key == 'TotalTime': assert int(new.text) >= int(old.text), 'editing time'
                        else: assert datetime.datetime.fromisoformat(new.text.replace('Z', '+00:00')) >= datetime.datetime.fromisoformat(old.text.replace('Z', '+00:00'))
                        metadata.append({'part': part, 'property': key, 'before': old.text, 'after': new.text}); old.text = new.text
                assert append.tree(left) == append.tree(right), 'metadata outside admitted fields'
                continue
            if part not in changed:
                assert before == after, 'unrelated part changed: ' + part
                continue
            row = changed[part]
            left, right = ET.fromstring(before), ET.fromstring(after)
            frame = left.find('p:cSld/p:spTree/p:graphicFrame', NS)
            actual_frame = right.find('p:cSld/p:spTree/p:graphicFrame', NS)
            para = frame.find('.//a:tbl/a:tr/a:tc/a:txBody/a:p', NS)
            assert len(para.findall('a:r', NS)) == 2
            for props in [r.find('a:rPr', NS) for r in para.findall('a:r', NS)] + [para.find('a:endParaRPr', NS)]:
                assert 'b' not in props.attrib and 'i' not in props.attrib
                props.attrib.update(STAGES[stage])
            location = 'p:nvGraphicFramePr/p:nvPr/p:extLst/p:ext/{' + append.P14 + '}modId'
            old_mod, actual_mod = frame.find(location, NS), actual_frame.find(location, NS)
            assert actual_mod is not None and set(actual_mod.attrib) == {'val'} and len(actual_mod) == 0
            value = actual_mod.attrib['val']; assert value.isdigit() and 0 <= int(value) <= 2**32 - 1 and value != old_mod.attrib['val']
            old_mod.set('val', value)
            assert append.tree(left) == append.tree(right), 'edited slide outside named style/modId changes: ' + part
            actual_para = actual_frame.find('.//a:tbl/a:tr/a:tc/a:txBody/a:p', NS)
            rows.append({'id': row['id'], 'part': part, 'slide': row['slide'], 'modId': value, 'runs': [{'text': r.find('a:t', NS).text, 'attributes': dict(r.find('a:rPr', NS).attrib)} for r in actual_para.findall('a:r', NS)], 'insertionAttributes': dict(actual_para.find('a:endParaRPr', NS).attrib), 'selector': row['selector'], 'declared': row['declared'], 'localFontOverrides': []})
    return {'sha256': append.sha(file), 'controls': sorted(rows, key=lambda r: r['slide']), 'metadata': metadata, 'unrelatedPartsByteExact': True}


def named_face(control, expected, stage):
    faces = {f['family']: f['postscriptName'] for f in expected['fonts']}
    regular = append.baseline.observed_face(control, faces)
    if stage == 'bold': return BOLD[regular], regular in SYNTHETIC_BOLD
    if stage == 'italic' and control['id'] == 'collection-he': return 'Arial-ItalicMT', False
    return regular, False


def inspect_pdf(file, directory, expected, stage, previous=None):
    assert append.sha(directory / 'baseline.pdf') == expected['styleBaseline']['pdf']
    rows = []
    with fitz.open(file) as doc, fitz.open(directory / 'baseline.pdf') as source:
        prior = fitz.open(previous) if previous else None
        try:
            assert len(doc) == len(source) == len(expected['controls'])
            for index, (page, old, control) in enumerate(zip(doc, source, expected['controls']), 1):
                assert page.rect == old.rect and abs(page.rect.width - 720) < .1 and abs(page.rect.height - 260) < .1
                label = ''.join(s['text'] for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['bbox'][1] < 60).strip()
                assert label == control['id'], 'PDF order/label'
                append.baseline.mixed.font_resources(doc, page)
                traces, before = append.body_traces(page), append.body_traces(old)
                assert traces and all(c[1] != 0 for t in traces for c in t['chars']), 'missing glyph'
                edited = index in expected['editTargets']; matrices, synthetic = [], False
                if edited:
                    face, synthetic = named_face(control, expected, stage)
                    omega_face = {'bold': 'Georgia-Bold', 'italic': 'Georgia-Italic', 'regular': 'Georgia'}[stage]
                    assert len(traces) == (3 if synthetic else 2), 'draw count'
                    omega = [t for t in traces if any(c[0] == 937 for c in t['chars'])]
                    original = [t for t in traces if all(c[0] != 937 for c in t['chars'])]
                    assert len(omega) == 1 and omega[0]['font'] == omega_face and omega[0]['type'] == 0 and len(omega[0]['chars']) == 1, 'Greek actual face'
                    assert all(t['font'] == face for t in original), 'original actual face'
                    assert [t['type'] for t in original] == ([0, 1] if synthetic else [0]), 'fill/stroke mode'
                    if synthetic:
                        assert original[0]['chars'] == original[1]['chars'], 'synthetic stroke glyphs/origins differ'
                        assert original[1]['linewidth'] == 4.5, 'synthetic stroke width'
                    source_text = ''.join(chr(c[0]) for t in before if t['type'] == 0 for c in t['chars'])
                    actual_text = ''.join(chr(c[0]) for t in traces if t['type'] == 0 for c in t['chars'])
                    assert actual_text == source_text, 'ToUnicode trace changed'
                    resources = {font[4]: font[3].split('+')[-1] for font in page.get_fonts()}
                    stream = b''.join(doc.xref_stream(x) for x in page.get_contents()).decode('latin1')
                    for match in re.findall(MATRIX, stream):
                        values = [float(v) for v in match[:6]]
                        if values[0] != 200: continue
                        resource_face = resources[match[6]]
                        assert resource_face in [face, omega_face] and values[1] == 0, 'unexpected body matrix'
                        original_matrix = resource_face == face
                        shear = stage == 'italic' and original_matrix and face != 'Arial-ItalicMT'
                        assert values[2] == (67.96875 if shear else 0), 'synthetic italic shear'
                        assert values[3] == (196.0784 if original_matrix and synthetic else 200), 'body vertical scale'
                        assert values[5] == (1085.206 if original_matrix and synthetic else 1083), 'body baseline matrix'
                        matrices.append({'face': resource_face, 'matrix': values, 'syntheticShear': shear})
                    assert len(matrices) == 2 and {m['face'] for m in matrices} == {face, omega_face}, 'missing body matrix'
                else: assert traces == before, 'unmodified trace'
                pix = page.get_pixmap(alpha=False)
                if not edited or stage == 'regular': assert pix.samples == old.get_pixmap(alpha=False).samples, 'baseline raster not restored/preserved'
                if prior: assert pix.samples == prior[index-1].get_pixmap(alpha=False).samples, 'stage round raster'
                image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples); visible = []
                for trace in traces:
                    text = ''.join(chr(c[0]) for c in trace['chars'])
                    if not text.strip(): continue
                    b = trace['bbox']; bounds = (max(0, math.floor(b[0])), max(0, math.floor(b[1])), min(720, math.ceil(b[2])), min(260, math.ceil(b[3])))
                    pixels = image.crop(bounds).tobytes(); ink = sum(min(pixels[i:i+3]) < 200 for i in range(0, len(pixels), 3))
                    assert ink > 5, 'invisible trace'
                    visible.append({'font': trace['font'], 'text': text, 'type': trace['type'], 'linewidth': trace['linewidth'], 'darkPixels': ink, 'bounds': bounds})
                rows.append({'id': control['id'], 'edited': edited, 'sourceText': control['text'] + ('Ω' if edited else ''), 'pdfFillTraceText': ''.join(chr(c[0]) for t in traces if t['type'] == 0 for c in t['chars']), 'actualFonts': sorted(set(t['font'] for t in traces)), 'syntheticBold': synthetic, 'textMatrices': matrices, 'visibleTraceInk': visible, 'rasterSha256': hashlib.sha256(pix.samples).hexdigest(), 'matchesPreStyleRaster': not edited or stage == 'regular'})
        finally:
            if prior: prior.close()
    return {'sha256': append.sha(file), 'controls': rows, 'stageRoundRasterMatch': bool(previous)}


def inspect(directory):
    expected = json.loads((directory / 'expected.json').read_text()); verify_baseline(directory, expected)
    stages = []
    for stage in STAGES:
        rounds = []
        for number in range(1, 4):
            stem = stage + '-round' + str(number)
            package = inspect_package(directory / (stem + '.pptx'), directory, expected, stage)
            if rounds: assert package['controls'] == rounds[0]['package']['controls'], 'stage source stability'
            pdf = inspect_pdf(directory / (stem + '.pdf'), directory, expected, stage, directory / (stage + '-round1.pdf') if number > 1 else None)
            rounds.append({'name': stem, 'package': package, 'pdf': pdf})
        with fitz.open(directory / (stage + '-round1.pdf')) as doc:
            sheet = Image.new('RGB', (720*3, 260*math.ceil(len(doc)/3)), 'white')
            for i, page in enumerate(doc):
                pix = page.get_pixmap(alpha=False); sheet.paste(Image.frombytes('RGB', (pix.width,pix.height), pix.samples), ((i%3)*720,(i//3)*260))
            sheet.save(directory / (stage + '-contact-sheet.png'))
        stages.append({'stage': stage, 'rounds': rounds})
    result = {'schemaVersion': 1, 'directory': str(directory), 'baseline': expected['styleBaseline'], 'editTargets': expected['editTargets'], 'stages': stages, 'scope': 'Named native bold, pure italic, restored regular. No theme/notes/Windows/source style writer or complete bidi/mark-outline claim.'}
    (directory / 'style-inspection.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    return result


if __name__ == '__main__':
    result = inspect(Path(sys.argv[1]))
    print(json.dumps({'editedControls': len(result['editTargets']), 'stages': len(result['stages']), 'nativeRoundsPerStage': 3}))
