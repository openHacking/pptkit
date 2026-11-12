import importlib.util
from pathlib import Path
import tempfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('replacement', Path(__file__).with_name('verify-corpus-image-replacement.py'))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
SOURCE = ROOT / 'tests/fixtures/corpus-image-replacement-source.pptx'
OUTPUT = ROOT / 'tests/fixtures/corpus-image-replacement-powerpoint.pptx'
MARKER = ROOT / 'tests/fixtures/corpus-image-replacement-marker.png'

class ReplacementEvidenceTests(unittest.TestCase):
    def test_retained_evidence(self):
        result = module.inspect(SOURCE, OUTPUT, MARKER)
        self.assertEqual(result['saved']['image-stretch']['sha256'], module.sha(MARKER.read_bytes()))

    def test_rejects_corrupted_replacement_control_and_geometry(self):
        for part, old, new in [
            ('ppt/media/image1.png', None, b'invalid replacement'),
            ('ppt/media/image2.png', None, b'invalid control'),
            ('ppt/slides/slide1.xml', b'cx="2235200"', b'cx="2235201"'),
        ]:
            with self.subTest(part=part), tempfile.TemporaryDirectory() as directory:
                altered = Path(directory) / 'altered.pptx'
                with zipfile.ZipFile(OUTPUT) as source, zipfile.ZipFile(altered, 'w') as target:
                    for info in source.infolist():
                        data = source.read(info.filename)
                        if info.filename == part:
                            if old is None: data = new
                            else:
                                self.assertIn(old, data)
                                data = data.replace(old, new, 1)
                        target.writestr(info, data)
                with self.assertRaises(ValueError): module.inspect(SOURCE, altered, MARKER)

if __name__ == '__main__': unittest.main()
