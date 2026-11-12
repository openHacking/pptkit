"""Recheck a captured notes experiment without rewriting its evidence manifest."""
import importlib.util
import json
import sys
from pathlib import Path

here = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('native_notes_pages', here / 'native-notes-pages.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

directory = Path(sys.argv[1]).resolve()
manifest = Path(sys.argv[2]).resolve()
actual = json.loads(json.dumps(module.inspect(directory)))
expected = json.loads(manifest.read_text())
for key in ('directory', 'application', 'files', 'rectOnlyChangeVerified', 'firstOpenPages',
            'rounds', 'copiedSlideRounds', 'rejectedControls', 'total', 'scope'):
    if actual[key] != expected[key]:
        raise AssertionError(f'Captured notes evidence mismatch: {key}')
print(json.dumps(actual['total']))
