import importlib.util
import json
import shutil
import tempfile
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location('surfaces', Path(__file__).with_name('verify-corpus-text-surfaces.py'))
surfaces = importlib.util.module_from_spec(spec)
spec.loader.exec_module(surfaces)


class SurfaceEvidenceControls(unittest.TestCase):
    def test_tampering_and_forged_capture_hashes_are_rejected(self):
        record = json.loads(surfaces.RECORD.read_text())
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            directory = root / record['artifactDirectory']
            directory.mkdir(parents=True)
            for name in record['files']:
                shutil.copy2(surfaces.ROOT / record['artifactDirectory'] / name, directory / name)
            for rel in ('tests/fixtures/corpus-text-powerpoint-interactive.pptx',
                        'scripts/research/verify-corpus-text-powerpoint.py'):
                (root / rel).parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(surfaces.ROOT / rel, root / rel)
            record_path = root / 'record.json'
            record_path.write_text(json.dumps(record))
            surfaces.verify(record_path, root)
            for name in ('text-notes.ax.txt', 'link-before.ax.txt', 'link-destination.ax.txt'):
                with self.subTest(name=name):
                    path = directory / name
                    original = path.read_bytes()
                    path.write_text('changed capture')
                    with self.assertRaisesRegex(ValueError, 'hash mismatch'):
                        surfaces.verify(record_path, root)
                    changed = dict(record, files=dict(record['files']))
                    changed['files'][name] = surfaces.digest(path)
                    record_path.write_text(json.dumps(changed))
                    with self.assertRaisesRegex(ValueError, 'capture'):
                        surfaces.verify(record_path, root)
                    path.write_bytes(original)
                    record_path.write_text(json.dumps(record))


if __name__ == '__main__':
    unittest.main()
