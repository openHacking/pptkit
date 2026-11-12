"""Recheck captured notes domains without modifying the capture or its manifest."""
import importlib.util
import json
import shutil
import sys
import tempfile
from pathlib import Path

source = Path(__file__).resolve().with_name('notes-image-domains.py')
spec = importlib.util.spec_from_file_location('notes_domains_recheck', source)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
capture_directory = Path(sys.argv[1]).resolve()
manifest = Path(sys.argv[2]).resolve()
expected = json.loads(manifest.read_text())
result = {}
write_text = Path.write_text


def capture_manifest(path, text, *args, **kwargs):
    if str(path) == 'docs/testing/notes-image-domains-evidence.json':
        result.update(json.loads(text))
        return len(text)
    return write_text(path, text, *args, **kwargs)


with tempfile.TemporaryDirectory(prefix='pptkit-notes-domain-recheck-') as temporary:
    isolated = Path(temporary) / 'capture'
    shutil.copytree(capture_directory, isolated)
    Path.write_text = capture_manifest
    try:
        module.inspect(isolated)
    finally:
        Path.write_text = write_text

assert result['total'] == expected['total']
assert len(result['groups']) == len(expected['groups'])
for actual, recorded in zip(result['groups'], expected['groups'], strict=True):
    for key in ('name', 'expectedSha256', 'sourceSha256', 'sourcePdfSha256',
                'rectSha256', 'sourceDomains', 'firstOpenPages', 'rounds', 'rejectedControls'):
        assert actual[key] == recorded[key], (actual['name'], key)
    # Recreated ZIPs can have different timestamps; compare their names and rejection reasons.
    assert [(item['name'], item['reason']) for item in actual['actualPptxCorruptionsRejected']] == [
        (item['name'], item['reason']) for item in recorded['actualPptxCorruptionsRejected']]
print('B/override semantic results and rejection reasons match the retained manifest.')
