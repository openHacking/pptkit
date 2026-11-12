"""Independent format metadata check layered on the unchanged six-page artwork oracle."""
import copy
import importlib.util
import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as E
HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('format_artwork_oracle', HERE/'master-artwork-native.py')
a = importlib.util.module_from_spec(spec)
spec.loader.exec_module(a)
EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'

def format_check(parts, native=False):
    node = a.n.root(parts, 'docProps/app.xml').find('{'+EP+'}PresentationFormat')
    assert node is not None and node.text == ('全屏显示(4:3)' if native else 'On-screen Show (4:3)'), 'Wrong or missing format label'
    size = a.n.root(parts, 'ppt/presentation.xml').find('p:sldSz', a.NS)
    assert size.get('cx') == '9144000' and size.get('cy') == '6858000', 'Format/frame disagreement'

def main():
    manifest = json.loads(Path('docs/testing/presentation-format-evidence.json').read_text())
    directory = Path(manifest['nativeDirectory'])
    source = directory/'input.pptx'
    assert a.sha(source.read_bytes()) == manifest['sourceSha256']
    original = a.n.payload(source)
    old = a.n.payload(Path(manifest['oldSourceFile']))
    assert a.sha(Path(manifest['oldSourceFile']).read_bytes()) == manifest['oldSourceSha256']
    assert original.keys() == old.keys()
    assert [p for p in original if original[p] != old[p]] == ['docProps/app.xml']
    assert original['docProps/app.xml'] == old['docProps/app.xml'].replace(b'On-screen Show (16:9)', b'On-screen Show (4:3)')
    format_check(original)
    expected, _ = a.semantic(original)
    a.package(source, expected)
    opened = (directory/'first-open.ax.txt').read_text()
    assert 'URL: file://'+str(source) in opened and not re.search('发现.*内容有问题|修复', opened)
    first = a.pdf(directory/'first.pdf', Path(manifest['oldSourceFile']).parent/'first.pdf')
    rounds = []
    baseline_meta = None
    for i in [1, 2, 3]:
        file = directory/f'format-round{i}.pptx'
        parts, _ = a.package(file, expected)
        format_check(parts, True)
        meta = a.metadata(parts)
        if baseline_meta is None: baseline_meta = meta
        else: a.same(meta, baseline_meta)
        pages = a.pdf(directory/f'format-round{i}.pdf', directory/'first.pdf')
        for phase in ['saved', 'reopened']:
            ax = (directory/f'format-round{i}-{phase}.ax.txt').read_text()
            assert 'URL: file://'+str(file) in ax and not re.search('发现.*内容有问题|修复', ax)
        assert str(file) not in (directory/f'format-round{i}-closed.ax.txt').read_text()
        rounds.append({'round':i, 'file':str(file), 'sha256':a.sha(file.read_bytes()), 'pages':pages, 'metadata':meta})
    assert '16.104 (25121423)' in (directory/'powerpoint-version.ax.txt').read_text()
    # Actual corrupt ZIPs demonstrate label checks independently of the body oracle.
    out = Path(tempfile.mkdtemp(prefix='negative-', dir=directory))
    negatives = []
    for name in ['wrong-label', 'missing-label', 'one-emu-frame', 'changed-body']:
        parts = copy.copy(original)
        if name in ['wrong-label', 'missing-label']:
            root = a.n.root(parts, 'docProps/app.xml')
            node = root.find('{'+EP+'}PresentationFormat')
            if name == 'wrong-label': node.text = 'On-screen Show (16:9)'
            else: root.remove(node)
            parts['docProps/app.xml'] = E.tostring(root, encoding='utf-8')
        elif name == 'one-emu-frame':
            root = a.n.root(parts, 'ppt/presentation.xml')
            root.find('p:sldSz', a.NS).set('cx', '9144001')
            parts['ppt/presentation.xml'] = E.tostring(root, encoding='utf-8')
        else:
            part = 'ppt/slides/slide1.xml'
            root = a.n.root(parts, part)
            next(root.iter('{'+a.A+'}t')).text = 'Wrong body'
            parts[part] = E.tostring(root, encoding='utf-8')
        file = out/(name+'.pptx')
        with zipfile.ZipFile(file, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for part, body in parts.items(): z.writestr(part, body)
        try:
            actual, _ = a.package(file, expected)
            format_check(actual)
        except AssertionError as error:
            negatives.append({'file':str(file), 'sha256':a.sha(file.read_bytes()), 'name':name, 'reason':str(error)})
        else: raise AssertionError('Accepted corrupted ZIP '+name)
    files = [{'file':str(f), 'bytes':f.stat().st_size, 'sha256':a.sha(f.read_bytes())} for f in sorted(directory.rglob('*')) if f.is_file()]
    dependencies = [{'file':str(f), 'sha256':a.sha(f.read_bytes())} for f in [Path(__file__), HERE/'master-artwork-native.py', HERE/'multi-master-native.py', HERE/'multi-master-operations.py']]
    result = {'schemaVersion':1, 'recordedAt':datetime.now(timezone.utc).isoformat(), 'verified':True, 'directory':str(directory), 'sourceSha256':manifest['sourceSha256'], 'application':'PowerPoint 16.104 (25121423)', 'historicalBaselineSha256':manifest['oldSourceSha256'], 'sourceChangedParts':['docProps/app.xml'], 'firstPages':first, 'rounds':rounds, 'negatives':negatives, 'dependencies':dependencies, 'artifacts':files, 'scope':'Metadata-only corrected 4:3 source. First print exact against historical artwork print; three independent save-close-exact-reopen rounds, 18 pages with strict geometry/text/raster/drawing/image/font comparison. Eight format scenes have automated consumers, not eight native runs. Existing artwork oracle normalizations unchanged; broader master operations and import writes unaccepted.'}
    Path('docs/testing/presentation-format-native-evidence.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n')
    print(json.dumps({'verified':True, 'firstPages':len(first), 'reopenedPages':sum(len(r['pages']) for r in rounds), 'negatives':len(negatives), 'artifacts':len(files)}))

if __name__ == '__main__': main()
