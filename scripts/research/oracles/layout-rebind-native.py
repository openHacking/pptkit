"""Independent actual rebind controls. This is not a general PDF font oracle."""
import copy
import importlib.util
import json
import re
import struct
import tempfile
from datetime import datetime, timezone
from pathlib import Path
import fitz

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('rebind_cell_primitives', HERE/'master-cell-properties-native.py')
c = importlib.util.module_from_spec(spec)
spec.loader.exec_module(c)
a, m = c.a, c.a.m


def font_pairs(document, font):
    """Admit only the observed simple native bfrange + format-6 TrueType map."""
    kind, value = document.xref_get_key(font[0], 'ToUnicode')
    assert kind == 'xref'
    cmap = document.xref_stream(int(value.split()[0])).decode('ascii')
    assert 'beginbfchar' not in cmap and 'usecmap' not in cmap
    codes = {}
    blocks = re.findall(r'(\d+) beginbfrange\s*(.*?)\s*endbfrange', cmap, re.S)
    assert blocks
    for count, body in blocks:
        lines = body.splitlines()
        assert len(lines) == int(count)
        for line in lines:
            match = re.fullmatch(r'<([0-9a-fA-F]{2})><([0-9a-fA-F]{2})><([0-9a-fA-F]{4})>', line.strip())
            assert match, 'Unknown native ToUnicode mapping'
            first, last, unicode = (int(s, 16) for s in match.groups())
            assert first <= last
            for code in range(first, last+1):
                assert code not in codes
                codes[code] = unicode+code-first
    data = document.extract_font(font[0])[3]
    tables = {}
    for i in range(struct.unpack_from('>H', data, 4)[0]):
        tag, _, start, length = struct.unpack_from('>4sIII', data, 12+16*i)
        assert start+length <= len(data)
        tables[tag] = data[start:start+length]
    table = tables[b'cmap']
    version, count = struct.unpack_from('>HH', table)
    assert version == 0 and count == 1, 'Unexpected cmap resource structure'
    _, _, offset = struct.unpack_from('>HHI', table, 4)
    format, length, language, first, total = struct.unpack_from('>HHHHH', table, offset)
    assert format == 6 and length == 10+total*2 and language == 0
    assert table[offset+length:] == b'\0\0', 'Unexpected native cmap trailer'
    pairs = {(unicode, struct.unpack_from('>H', table, offset+10+2*(code-first))[0]) for code, unicode in codes.items() if first <= code < first+total}
    assert len(pairs) == len(codes)
    return pairs, data


def glyph_trace(document, page):
    catalog = [(font[3].split('+')[-1], *font_pairs(document, font)) for font in page.get_fonts()]
    result = copy.deepcopy(page.get_texttrace())
    for span in result:
        needed = {(char[0], char[1]) for char in span['chars']}
        candidates = [data for face, pairs, data in catalog if face == span['font'] and needed <= pairs]
        assert candidates, 'No matching actual PDF font resource'
        alternatives = [tuple((char[0], m.glyph_program(data, char[1]), char[2], char[3]) for char in span['chars']) for data in candidates]
        assert all(value == alternatives[0] for value in alternatives), 'Ambiguous font resources differ in glyph metrics or outlines'
        span['chars'] = alternatives[0]
    return a.plain(result)


def page_records(file):
    document = fitz.open(file)
    assert len(document) == 3, 'Hidden slide omitted or unexpected page count'
    return [a.pdf_page_record(document, i) for i in range(3)]


def main():
    manifest = json.loads(Path('docs/testing/layout-rebind-evidence.json').read_text())
    directory = Path(manifest['nativeDirectory'])
    old = fitz.open(directory/'before.pdf')
    new = fitz.open(directory/'first.pdf')
    assert len(old) == len(new) == 3
    # Page 2 is untouched, including raster/drawings/images and all used glyphs.
    for i in [0, 2]:
        assert glyph_trace(old, old[i]) == glyph_trace(new, new[i]), ('Changed foreground glyphs or line metrics', i)
        before, after = a.pdf_page_record(old, i), a.pdf_page_record(new, i)
        for key in ['page', 'size', 'text', 'images']:
            assert before[key] == after[key], ('Changed bystander content', i, key)
        if i == 2:
            for key in ['rasterSha256', 'drawings']:
                assert before[key] == after[key], ('Changed unselected page', key)
        else:
            changes = [(j, x, y) for j, (x, y) in enumerate(zip(before['drawings'], after['drawings'])) if x != y]
            assert len(before['drawings']) == len(after['drawings']) and len(changes) == 1
            _, x, y = changes[0]
            assert set(k for k in x if x[k] != y[k]) == {'fill'}, 'Unexpected preserved-page drawing change'
            # Exact observed native PDF components for authored CC2200/0044CC.
            assert x['fill'] == [0.800000011920929, 0.13331808149814606, 0.0]
            assert y['fill'] == [0.0, 0.2666666805744171, 0.800000011920929], 'Wrong new owner artwork'
    first = page_records(directory/'first.pdf')
    for n in range(1, 4):
        assert page_records(directory/f'rebind-round{n}.pdf') == first, ('Native reopen changed PDF content', n)
    parts = [a.n.payload(directory/f'rebind-round{n}.pptx') for n in range(1, 4)]
    assert c.full(parts[0]) == c.full(parts[1]) == c.full(parts[2]), 'Native package changed after first save'
    # Real failure controls are retained; this comparison must reject the old
    # three-cell baseline drift without any coordinate/size rounding.
    failed = Path('.tmp/layout-rebind/native-0OWvU8')
    failed_before = fitz.open(failed/'before.pdf')
    failed_after = fitz.open(failed/'first.pdf')
    assert page_records(failed/'before.pdf') == page_records(directory/'before.pdf'), 'Paragraph-font fix changed the source native baseline'
    assert glyph_trace(failed_before, failed_before[0]) != glyph_trace(failed_after, failed_after[0])
    for name in ['font-all', 'font-latin', 'font-aux']:
        control = fitz.open(failed/(name+'.pdf'))
        assert glyph_trace(control, control[0]) == glyph_trace(failed_before, failed_before[0])
    for name in ['end-font', 'default-font']:
        control = fitz.open(failed/(name+'.pdf'))
        assert glyph_trace(control, control[0]) == glyph_trace(failed_before, failed_before[0])
    out = Path(tempfile.mkdtemp(prefix='typography-negatives-', dir=directory))
    negatives = []
    for name in ['covered-text', 'swapped-pages', 'missing-hidden-page', 'same-raster-font-advance']:
        doc = fitz.open(directory/'first.pdf')
        if name == 'covered-text':
            doc[2].draw_rect(doc[2].search_for('Shared')[0], color=(1, 1, 1), fill=(1, 1, 1))
        elif name == 'swapped-pages':
            doc.select([2, 1, 0])
        elif name == 'missing-hidden-page':
            doc.select([0, 2])
        else:
            span = doc[2].get_texttrace()[0]
            needed = {(char[0], char[1]) for char in span['chars']}
            candidates = [font_pairs(doc, font)[1] for font in doc[2].get_fonts() if font[3].split('+')[-1] == span['font'] and needed <= font_pairs(doc, font)[0]]
            assert len(candidates) == 1
            program = candidates[0]
            gid = span['chars'][0][1]
            changed = m.alter_font_advance(program, gid)
            before, after = m.glyph_program(program, gid), m.glyph_program(changed, gid)
            assert after == (before[0], before[1]+1, before[2], before[3])
            streams = [xref for xref in range(1, doc.xref_length()) if doc.xref_is_stream(xref) and doc.xref_stream(xref) == program]
            assert len(streams) == 1
            doc.update_stream(streams[0], changed)
        file = out/(name+'.pdf')
        doc.save(file)
        try:
            assert page_records(file) == first, 'Changed actual native PDF records'
        except AssertionError as error:
            reason = str(error)
        else:
            raise AssertionError('Accepted corrupt actual PDF '+name)
        if name == 'same-raster-font-advance':
            patched = fitz.open(file)
            assert all(patched[i].get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False).samples == new[i].get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False).samples for i in range(3))
            assert glyph_trace(patched, patched[2]) != glyph_trace(new, new[2]), 'Missed font metric damage behind identical pixels'
        negatives.append({'name': name, 'file': str(file), 'bytes': file.stat().st_size, 'sha256': a.sha(file.read_bytes()), 'reason': reason})
    for n in range(1, 4):
        reopened = (directory/f'rebind-round{n}-reopened.ax.txt').read_text()
        assert 'URL: file://'+str(directory/f'rebind-round{n}.pptx') in reopened
        assert not re.search('发现.*内容有问题|修复', reopened)
    for name in ['before', 'first', 'rebind-round1', 'rebind-round2', 'rebind-round3']:
        options = (directory/(name+'-print-options.ax.txt')).read_text()
        assert '复选框 打印隐藏幻灯片, Value: 1' in options and '页码：3/3' in options
    assert a.sha((directory/'input.pptx').read_bytes()) == manifest['sourceSha256']
    assert a.sha((directory/'before.pptx').read_bytes()) == manifest['beforeSha256']
    dependencies = [Path(__file__), HERE/'master-cell-properties-native.py', HERE/'master-artwork-native.py', HERE/'multi-master-operations.py', HERE/'multi-master-native.py']
    artifacts = [{'file': str(file), 'bytes': file.stat().st_size, 'sha256': a.sha(file.read_bytes())} for root in [directory, failed] for file in sorted(root.iterdir()) if file.is_file() and not file.name.startswith('~$')]
    result = {'schemaVersion': 1, 'recordedAt': datetime.now(timezone.utc).isoformat(), 'verifiedTypographyAndReopen': True, 'completeNativeRebindAccepted': False, 'directory': str(directory), 'failedDirectory': str(failed), 'sourceSha256': manifest['sourceSha256'], 'beforeSha256': manifest['beforeSha256'], 'pages': 3, 'rounds': 3, 'glyphBystanderPages': [0, 2], 'negatives': negatives, 'artifacts': artifacts, 'dependencies': [{'file': str(file), 'sha256': a.sha(file.read_bytes())} for file in dependencies], 'scope': 'Exact used ASCII glyph programs/metrics, text positions, images and drawing geometry on preserve/unselected pages; only the named master accent fill changes. Entire actual PDF records exact through three reopen rounds; every native package part exact after first save except validated clocks/TotalTime. Captured first-save additions/defaults, reset-page semantics, complete source ownership/chart/workbook graph and Parser no-op binding remain under review. This bounded proof does not accept the full rebind API or all first-save deltas.'}
    Path('docs/testing/layout-rebind-native-typography-evidence.json').write_text(json.dumps(result, indent=2)+'\n')
    print(json.dumps({'verifiedTypographyAndReopen': True, 'completeNativeRebindAccepted': False, 'negatives': len(negatives)}))


if __name__ == '__main__':
    main()
