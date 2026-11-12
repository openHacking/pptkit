"""Captured PowerPoint append contract; separate from the strict no-edit oracle."""
import copy, hashlib, importlib.util, json, math, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image, ImageDraw

spec = importlib.util.spec_from_file_location('notes', Path(__file__).with_name('native-notes-fonts.py'))
notes = importlib.util.module_from_spec(spec)
spec.loader.exec_module(notes)
NS, A = notes.NS, notes.A
TARGETS = [3, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15]
ANCILLARY = {'ppt/viewProps.xml', 'docProps/app.xml', 'docProps/core.xml'}

def paragraph(n):
    ps = notes.body(n).findall('p:txBody/a:p', NS)
    assert len(ps) == 1
    return ps[0]

def expected_note(before, control, index):
    n = copy.deepcopy(before)
    p = paragraph(n)
    runs = p.findall('a:r', NS)
    assert len(runs) == 1
    rp = runs[0].find('a:rPr', NS)
    assert rp.attrib == {'lang': control['language']}
    rp.set('dirty', '0')
    added = copy.deepcopy(runs[0])
    added.find('a:t', NS).text = 'Ω'
    arp = added.find('a:rPr', NS)
    arp.attrib.clear()
    arp.set('lang', 'en-US')
    if control['language'] not in ['ar', 'he']:
        arp.set('altLang', control['language'])
    if index != 15:
        arp.set('dirty', '0')
        arp.set('err', '1')
    p.insert(list(p).index(runs[0]) + 1, added)
    end = p.find('a:endParaRPr', NS)
    assert end.attrib == {'lang': 'en-US'}
    end.set('dirty', '0')
    return n

def inspect_package(file, directory, expected):
    assert notes.sha(directory/'baseline.pptx') == expected['nativeBaseline']['sha256']
    records, ancillary = [], {}
    with zipfile.ZipFile(directory/'baseline.pptx') as old, zipfile.ZipFile(file) as z:
        assert len(z.namelist()) == len(set(z.namelist()))
        assert set(z.namelist()) == set(old.namelist()), 'part inventory'
        changed = []
        for part in old.namelist():
            before, after = old.read(part), z.read(part)
            if before != after:
                changed.append(part)
            if part in ANCILLARY:
                ancillary[part] = {'baselineSha256': hashlib.sha256(before).hexdigest(),
                                   'nativeSha256': hashlib.sha256(after).hexdigest(),
                                   'baselineXml': before.decode(), 'nativeXml': after.decode()}
                continue
            if part.startswith('ppt/notesSlides/notesSlide') and part.endswith('.xml'):
                index = int(Path(part).stem.removeprefix('notesSlide'))
                target = ET.fromstring(before)
                if index in TARGETS:
                    target = expected_note(target, expected['controls'][index-1], index)
                actual = ET.fromstring(after)
                assert notes.tree(actual) == notes.tree(target), 'unexpected notes change: '+part
                records.append({'slide': index, 'semanticSha256': notes.digest(actual),
                                'runProperties': [ET.tostring(r.find('a:rPr', NS), encoding='unicode')
                                                  for r in paragraph(actual).findall('a:r', NS)]})
            else:
                assert after == before, 'non-notes payload changed: '+part
        assert len(records) == 15
    return {'sha256': notes.sha(file), 'controls': sorted(records, key=lambda r:r['slide']),
            'changedPayloads': changed, 'ancillaryRecords': ancillary,
            'scope': 'All ZIP payloads exact except complete edited notes trees and three recorded metadata/view parts.'}

def strip_sequence(trace):
    return {k:v for k,v in trace.items() if k != 'seqno'}

def inspect_pdf(file, directory, expected, reference=None):
    assert notes.sha(directory/'baseline-notes.pdf') == expected['nativeBaseline']['pdfSha256']
    rows = []
    with fitz.open(directory/'baseline-notes.pdf') as baseline, fitz.open(file) as doc:
        assert len(doc) == len(baseline) == 15
        ref = fitz.open(reference) if reference else None
        try:
            for i, (page, old) in enumerate(zip(doc, baseline), 1):
                assert tuple(page.rect) == tuple(old.rect) == (0, 0, 595, 842)
                traces, original = page.get_texttrace(), old.get_texttrace()
                appended = [t for t in traces if t['bbox'][1]>410 and ''.join(chr(c[0]) for c in t['chars'])=='Ω'] if i in TARGETS else []
                assert len(appended) == (1 if i in TARGETS else 0), 'appended trace count'
                kept = [t for t in traces if not any(t is a for a in appended)]
                assert [strip_sequence(t) for t in kept] == [strip_sequence(t) for t in original], 'original glyph/face/position or neighbor changed'
                pix, oldpix = page.get_pixmap(alpha=False), old.get_pixmap(alpha=False)
                image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
                oldimage = Image.frombytes('RGB', (oldpix.width, oldpix.height), oldpix.samples)
                ink = None
                if appended:
                    t = appended[0]
                    face = 'ArialMT' if expected['variant']=='A' or 10<=i<=12 else 'Georgia'
                    assert t['font'] == face and abs(t['size']-41)<.001 and t['type']==0, 'actual appended face/size/paint'
                    assert len(t['chars'])==1 and t['chars'][0][0]==937 and t['chars'][0][1]!=0
                    b = t['bbox']
                    bounds = (math.floor(b[0])-1, math.floor(b[1])-1, math.ceil(b[2])+1, math.ceil(b[3])+1)
                    pixels = image.crop(bounds).tobytes()
                    ink = sum(min(pixels[j:j+3])<200 for j in range(0,len(pixels),3))
                    assert ink>20, 'extractable but invisible append'
                    ImageDraw.Draw(image).rectangle(bounds, fill='white')
                    ImageDraw.Draw(oldimage).rectangle(bounds, fill='white')
                assert image.tobytes()==oldimage.tobytes(), 'raster changed outside appended glyph'
                notes.table.mixed.font_resources(doc, page)
                if ref:
                    assert traces == ref[i-1].get_texttrace(), 'later round trace drift'
                    assert pix.samples == ref[i-1].get_pixmap(alpha=False).samples, 'later round raster drift'
                rows.append({'slide':i, 'appended':bool(appended), 'actualFont': appended[0]['font'] if appended else None,
                             'appendTrace': appended, 'visibleAppendInk':ink,
                             'originalTraceAndOutsideRasterExact':True, 'rasterSha256':hashlib.sha256(pix.samples).hexdigest()})
        finally:
            if ref:
                ref.close()
    return {'sha256': notes.sha(file), 'pages':rows, 'laterRoundMatchesFirst':bool(reference)}

def inspect(directory):
    expected=json.loads((directory/'expected.json').read_text())
    assert expected['editTargets']==TARGETS
    rounds=[]
    for number in [1,2,3]:
        name='append-round'+str(number)
        package=inspect_package(directory/(name+'.pptx'),directory,expected)
        if rounds:
            assert package['controls']==rounds[0]['package']['controls'], 'native notes tree drift'
        pdf=inspect_pdf(directory/(name+'-notes.pdf'),directory,expected,
                        directory/'append-round1-notes.pdf' if number>1 else None)
        rounds.append({'name':name,'package':package,'pdf':pdf})
    result={'variant':expected['variant'],'rounds':rounds,'nativeReopenedNotesPages':45,
            'limits':'Append only. No local styles/theme changes, complete notes artwork, source style writer or broad font equivalence.'}
    (directory/'append-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    return result

if __name__=='__main__':
    result=inspect(Path(sys.argv[1]))
    print(json.dumps({'variant':result['variant'],'nativeRounds':3,'pages':45,'appends':33}))
