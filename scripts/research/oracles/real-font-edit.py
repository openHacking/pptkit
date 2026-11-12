"""Named PowerPoint SC edit/style/theme experiment. No production admission rules."""
import copy, hashlib, importlib.util, json, posixpath, re, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image
spec = importlib.util.spec_from_file_location('fresh', Path(__file__).with_name('real-font-fresh.py'))
fresh = importlib.util.module_from_spec(spec); spec.loader.exec_module(fresh)
NS = fresh.NS
STAGES = {'edit': (False, False, 'HiraginoSansGB-W3'), 'bold': (True, False, 'HiraginoSansGB-W6'), 'italic': (False, True, 'HiraginoSansGB-W3'), 'theme': (False, False, 'HiraginoSansGB-W3')}

def relationship(z, owner, relation_type, relation_id=None):
    file = posixpath.join(posixpath.dirname(owner), '_rels', posixpath.basename(owner)+'.rels')
    rels = [r for r in ET.fromstring(z.read(file)) if r.attrib['Type'].endswith('/'+relation_type) and (relation_id is None or r.attrib['Id'] == relation_id)]
    assert len(rels) == 1 and rels[0].attrib.get('TargetMode') != 'External', 'ambiguous/external relationship'
    target = rels[0].attrib['Target']
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))

def read_package(file, expected_count=16):
    rows = []
    with zipfile.ZipFile(file) as z:
        assert len(z.namelist()) == len(set(z.namelist()))
        presentation = ET.fromstring(z.read('ppt/presentation.xml'))
        for slide in presentation.findall('p:sldIdLst/p:sldId', NS):
            part = relationship(z, 'ppt/presentation.xml', 'slide', slide.attrib['{'+NS['r']+'}id'])
            root = ET.fromstring(z.read(part)); shapes = {}
            for s in root.findall('p:cSld/p:spTree/p:sp', NS):
                name = s.find('p:nvSpPr/p:cNvPr', NS).attrib['name']; assert name not in shapes
                shapes[name] = s
            assert set(shapes) == {'Label', 'Probe'}
            body = shapes['Probe'].find('p:txBody', NS)
            assert body.find('a:bodyPr', NS).attrib['wrap'] == 'none'
            runs = []
            for run in body.findall('a:p/a:r', NS):
                style = run.find('a:rPr', NS)
                slots = [style.find('a:'+s, NS).attrib['typeface'] for s in ['latin','ea','cs']]
                runs.append({'text': run.find('a:t', NS).text, 'language': style.attrib.get('lang'), 'altLanguage': style.attrib.get('altLang'), 'size': style.attrib.get('sz'), 'bold': style.attrib.get('b','0') == '1', 'italic': style.attrib.get('i','0') == '1', 'slots': slots, 'rawAttributes': dict(style.attrib)})
            layout = relationship(z, part, 'slideLayout'); master = relationship(z, layout, 'slideMaster'); theme = relationship(z, master, 'theme')
            data = z.read(theme); theme_root = ET.fromstring(data)
            rows.append({'id': ''.join(shapes['Label'].find('p:txBody',NS).itertext()), 'part': part, 'runs': runs, 'bodyProperties':dict(body.find('a:bodyPr',NS).attrib), 'paragraphProperties':[ET.tostring(q.find('a:pPr',NS),encoding='unicode') if q.find('a:pPr',NS) is not None else None for q in body.findall('a:p',NS)], 'geometry':ET.tostring(shapes['Probe'].find('p:spPr/a:xfrm',NS),encoding='unicode'), 'theme': {'part':theme, 'name':theme_root.attrib['name'], 'sha256':hashlib.sha256(data).hexdigest(), 'fontScheme': {'sha256':hashlib.sha256(ET.tostring(theme_root.find('a:themeElements/a:fontScheme',NS))).hexdigest(), 'majorLatin':theme_root.find('a:themeElements/a:fontScheme/a:majorFont/a:latin',NS).attrib['typeface'], 'minorLatin':theme_root.find('a:themeElements/a:fontScheme/a:minorFont/a:latin',NS).attrib['typeface']}}})
    assert len(rows) == expected_count
    return rows

def semantic_run(run): return {k:v for k,v in run.items() if k != 'rawAttributes'}
def inspect_package(file, source_rows, stage):
    bold, italic, _ = STAGES[stage]; rows = read_package(file)
    for row, source in zip(rows, source_rows):
        assert row['id'] == source['id'], 'slide order/label'
        assert row['geometry'] == source['geometry'] and row['bodyProperties'] == source['bodyProperties'], 'frame changed'
        assert row['paragraphProperties'] == ([None] if row['id'] == 'map-sc' else source['paragraphProperties']), 'paragraph properties differ'
        if row['id'] != 'map-sc':
            assert [semantic_run(r) for r in row['runs']] == [semantic_run(r) for r in source['runs']], 'unrelated text/font/style changed'
        else:
            expected = [{'text':'骨直令返', 'language':'zh-CN', 'altLanguage':None}, {'text':'Ω', 'language':'el-GR', 'altLanguage':'zh-CN'}]
            expected = [{**r,'size':'4800','bold':bold,'italic':italic,'slots':['Hiragino Sans GB']*3} for r in expected]
            assert [semantic_run(r) for r in row['runs']] == expected, 'edited topology/language/slots/style differ'
        if stage == 'theme':
            assert row['theme']['sha256'] != source['theme']['sha256'] and row['theme']['fontScheme'] != source['theme']['fontScheme'], 'theme did not change'
            assert row['theme']['name'].startswith('Office'), 'unexpected chosen theme'
        else:
            assert row['theme'] == source['theme'], 'unexpected theme change'
    return {'sha256':fresh.sha(file), 'controls':rows}

def inspect_pdf(file, expected, stage, baseline=None):
    _, italic, face = STAGES[stage]; doc = fitz.open(file); assert len(doc) == 16
    previous = fitz.open(baseline) if baseline else None
    fonts = {f['id']:f['postscriptName'] for f in expected['fonts']}; rows = []
    for index, (page, control) in enumerate(zip(doc, expected['controls'])):
        assert abs(page.rect.width-720)<.1 and abs(page.rect.height-260)<.1
        label = ''.join(s['text'] for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['bbox'][1]<60).strip()
        assert label == control['id'], 'PDF page order'
        traces = [t for t in page.get_texttrace() if t['bbox'][1]>60]; glyphs = [c for t in traces for c in t['chars']]
        actual = sorted(set(t['font'] for t in traces)); desired = face if control['id'] == 'map-sc' else fonts[control['fontId']]
        assert actual == [desired] and glyphs and all(c[1] != 0 for c in glyphs), 'actual face/glyph differs'
        if control['id'] == 'map-sc': assert ''.join(chr(c[0]) for c in glyphs) == '骨直令返Ω', 'PDF edited text differs'
        pix = page.get_pixmap(alpha=False); image = Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
        ink = sum(min(p) < 200 for p in image.crop((0,60,720,260)).getdata()); assert ink>100
        if previous:
            old = previous[index].get_pixmap(alpha=False)
            assert (pix.width,pix.height,pix.samples) == (old.width,old.height,old.samples), 'round raster differs'
        row = {'id':control['id'],'fonts':actual,'glyphCount':len(glyphs),'glyphZeroCount':0,'darkPixels':ink,'rasterSha256':hashlib.sha256(pix.samples).hexdigest()}
        if control['id'] == 'map-sc':
            # Narrowly inspect the named PowerPoint PDF body text matrices; not a general PDF interpreter.
            stream = b''.join(doc.xref_stream(x) for x in page.get_contents()).decode('latin1')
            matches = re.findall(r'BT\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm\s+/(C\d+)\s+1\s+Tf',stream)
            matrices = [[float(v) for v in m[:6]] for m in matches]
            assert matrices and all(m[0] == m[3] == 200 and m[1] == 0 for m in matrices)
            assert all(abs(m[2]/m[0] - (67.96875/200 if italic else 0)) < 1e-9 for m in matrices), 'body shear differs'
            row['bodyTextMatrices'] = matrices; row['syntheticShear'] = italic
        rows.append(row)
    return {'sha256':fresh.sha(file),'controls':rows,'baselineRasterMatch':bool(previous)}

def inspect(directory):
    expected = json.loads((directory/'expected.json').read_text()); fresh.inspect_package(directory/'fresh-fonts.pptx',expected,True)
    for font in expected['fonts']: assert fresh.sha(font['file']) == font['sha256']
    source = read_package(directory/'fresh-round3.pptx'); results = {}
    sheet = Image.new('RGB',(720,260*4),'white')
    original = fitz.open(directory/'fresh-fonts.pdf')
    for stage in STAGES:
        name = 'fresh-'+stage+'-sc'; pdf = directory/(name+'.pdf')
        results[stage] = {'initial': {'package':inspect_package(directory/(name+'.pptx'),source,stage),'pdf':inspect_pdf(pdf,expected,stage)},'rounds':[]}
        with fitz.open(pdf) as doc:
            pix = doc[1].get_pixmap(alpha=False); sheet.paste(Image.frombytes('RGB',(pix.width,pix.height),pix.samples),(0,list(STAGES).index(stage)*260))
            if stage != 'theme':
                for i in range(16):
                    if i != 1: assert doc[i].get_pixmap(alpha=False).samples == original[i].get_pixmap(alpha=False).samples, 'unrelated raster changed'
        for round in range(1,4):
            stem = name+'-round'+str(round)
            package = inspect_package(directory/(stem+'.pptx'),source,stage)
            wanted = copy.deepcopy(results[stage]['initial']['package']['controls'])
            if stage == 'edit':
                # Observed native proofing normalization: only the appended run gains dirty=0.
                assert 'dirty' not in wanted[1]['runs'][1]['rawAttributes']
                wanted[1]['runs'][1]['rawAttributes']['dirty'] = '0'
            assert package['controls'] == wanted, 'round raw topology/theme differs'
            results[stage]['rounds'].append({'name':stem,'package':package,'pdf':inspect_pdf(directory/(stem+'.pdf'),expected,stage,baseline=pdf)})
    # Style operations actually changed pixels; font and style declarations alone are insufficient.
    edited = results['edit']['initial']['pdf']['controls'][1]['rasterSha256']
    assert all(results[stage]['initial']['pdf']['controls'][1]['rasterSha256'] != edited for stage in ['bold','italic','theme'])
    sheet.save(directory/'edit-contact-sheet.png')
    return {'schemaVersion':1,'artifactsDirectory':str(directory),'application':'Microsoft PowerPoint 16.104 (16.104.25121423)','scope':'SC map materialization: append Greek, bold, synthetic italic, Office theme; three native save/close/reopen rounds per stage','sourceSha256':expected['sourceSha256'],'results':results}

if __name__ == '__main__':
    directory = Path(sys.argv[1]); result = inspect(directory)
    (directory/'edit-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'stages':list(STAGES),'roundsPerStage':3,'pagesPerArtifact':16,'directory':str(directory)}))
