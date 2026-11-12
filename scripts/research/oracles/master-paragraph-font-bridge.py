"""Current table-font producer bridge; historical native rounds stay historical."""
import importlib.util
import json
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import fitz
HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('font_bridge_artwork', HERE/'master-artwork-native.py')
a = importlib.util.module_from_spec(spec); spec.loader.exec_module(a)


def main():
    manifest = json.loads(Path('docs/testing/master-paragraph-font-bridge-evidence.json').read_text())
    directory = Path(manifest['directory'])
    version = (directory/'powerpoint-version.ax.txt').read_text()
    assert '版本 16.104 (25121423)' in version
    records = []
    for record in manifest['records']:
        for key, hash_key in [('file', 'sha256'), ('historicalFile', 'historicalSha256'), ('referencePdf', 'referencePdfSha256')]:
            assert a.sha(Path(record[key]).read_bytes()) == record[hash_key]
        mode = record['mode']
        opened = (directory/(mode+'-open.ax.txt')).read_text()
        assert 'URL: file://'+record['file'] in opened
        assert '发现内容有问题' not in opened and '修复' not in opened
        options = (directory/(mode+'-print-options.ax.txt')).read_text()
        assert '页码：4/4' in options and '要打印的版式, Value: 幻灯片' in options
        file = directory/(mode+'.pdf')
        current, old = fitz.open(file), fitz.open(record['referencePdf'])
        assert len(current) == len(old) == 4
        pages = [a.pdf_page_record(current, i) for i in range(4)]
        assert pages == [a.pdf_page_record(old, i) for i in range(4)], 'Current table-font output changed actual native page records'
        records.append({**record, 'pdf': str(file), 'pdfSha256': a.sha(file.read_bytes()), 'fullPageRecordsExact': True, 'pages': pages})
    assert [r['mode'] for r in records] == ['before', 'after']
    negative_directory = Path(tempfile.mkdtemp(prefix='negative-', dir=directory))
    negative_fixtures = []
    with zipfile.ZipFile(records[0]['file']) as archive: source = {name: archive.read(name) for name in archive.namelist()}
    font_mark = b'<a:endParaRPr lang="en-US"><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:endParaRPr>'
    for name, part, before, after in [
        ('missing-font-mark', 'ppt/slideMasters/slideMaster2.xml', font_mark, b'<a:endParaRPr lang="en-US"/>'),
        ('changed-run-size', 'ppt/slideMasters/slideMaster2.xml', b'sz="1400"', b'sz="1500"'),
        ('unhide-master-artwork', 'ppt/slides/slide2.xml', b'showMasterSp="0"', b'showMasterSp="1"'),
        ('changed-metadata', 'docProps/app.xml', b'<Application>PPTKit</Application>', b'<Application>Other</Application>'),
    ]:
        assert before in source[part]
        parts = dict(source); parts[part] = parts[part].replace(before, after, 1)
        file = negative_directory/(name+'.pptx')
        with zipfile.ZipFile(file, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            for key, body in parts.items(): archive.writestr(key, body)
        negative_fixtures.append({'name': name, 'file': str(file), 'sha256': a.sha(file.read_bytes()), 'bytes': file.stat().st_size, 'changedPart': part})
    artifacts = [{'file': str(file), 'bytes': file.stat().st_size, 'sha256': a.sha(file.read_bytes())} for file in sorted(directory.iterdir()) if file.is_file() and not file.name.startswith('~$')]
    deps = [Path(__file__), HERE/'master-artwork-native.py', HERE/'multi-master-native.py', HERE/'multi-master-operations.py']
    baselines = [Path('docs/testing/master-cell-properties-native-evidence.json'), Path('docs/testing/master-shared-native-evidence.json')]
    result = {'schemaVersion': 1, 'recordedAt': datetime.now(timezone.utc).isoformat(), 'verified': True, 'directory': str(directory), 'records': records, 'artifacts': artifacts, 'negativeFixtures': negative_fixtures, 'dependencies': [{'file': str(file), 'sha256': a.sha(file.read_bytes())} for file in deps], 'baselineEvidence': [{'file': str(file), 'sha256': a.sha(file.read_bytes())} for file in baselines], 'scope': 'Two actual current-producer first-open prints, before and after the shared heading edit, match all four corresponding historical native page records exactly: raster, text, used trace positions/glyph IDs, complete embedded font streams, drawings and images. Only six paragraph-mark font additions in two master-table payloads are admitted by the separate byte-delta verifier. Historical source hashes/native rounds remain unchanged; these are two new first-open controls, not new save/reopen rounds. Negative fixture bytes are recorded here; rejection is verified by the separate producer bridge verifier.'}
    Path('docs/testing/master-paragraph-font-native-evidence.json').write_text(json.dumps(result, indent=2)+'\n')
    print(json.dumps({'verified': True, 'actualFiles': 2, 'exactNativePages': 8}))


if __name__ == '__main__': main()
