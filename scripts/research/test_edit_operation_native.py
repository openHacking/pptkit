import importlib.util
from pathlib import Path
import tempfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('native_edit', Path(__file__).with_name('verify-edit-operation-native.py'))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
OUTPUT = ROOT / 'tests/fixtures/edit-operation-native-powerpoint.pptx'

class NativeEditEvidenceTests(unittest.TestCase):
    def test_retained(self): module.verify()

    def test_rejects_text_group_and_geometry_corruption(self):
        for old, new in [(b'>native<', b'>wrong<'), (b'name="group"', b'name="other"'), (b'x="2286000"', b'x="2286001"')]:
            with self.subTest(old=old), tempfile.TemporaryDirectory() as directory:
                target_path = Path(directory) / 'altered.pptx'
                with zipfile.ZipFile(OUTPUT) as source, zipfile.ZipFile(target_path, 'w') as target:
                    for info in source.infolist():
                        data = source.read(info.filename)
                        if info.filename == 'ppt/slides/slide1.xml':
                            self.assertIn(old, data)
                            data = data.replace(old, new, 1)
                        target.writestr(info, data)
                with self.assertRaises(ValueError): module.inspect(target_path, 'Edited native')

if __name__ == '__main__': unittest.main()
