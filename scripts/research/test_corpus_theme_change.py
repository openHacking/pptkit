import importlib.util
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

spec = importlib.util.spec_from_file_location('theme_change', Path(__file__).with_name('verify-corpus-theme-change.py'))
verify = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verify)


class ThemeChangeControls(unittest.TestCase):
    def test_corrupted_saved_package_is_rejected(self):
        source = verify.DIR / 'theme-layout.pptx'
        with ZipFile(source) as z:
            originals = {info.filename: z.read(info) for info in z.infolist()}
        cases = [
            ('ppt/theme/theme1.xml', b'name="Archway"', b'name="Other"'),
            ('ppt/slides/slide1.xml', b'Inherited title', b'Corrupted title'),
            ('ppt/slides/slide2.xml', b'show="0"', b'show="1"'),
            ('ppt/slides/_rels/slide1.xml.rels', b'slideLayout28.xml', b'slideLayout999.xml'),
        ]
        with tempfile.TemporaryDirectory() as temporary:
            for part, old, new in cases:
                with self.subTest(part=part):
                    self.assertIn(old, originals[part])
                    candidate = Path(temporary) / 'candidate.pptx'
                    with ZipFile(candidate, 'w') as out:
                        for name, contents in originals.items():
                            out.writestr(name, contents.replace(old, new, 1) if name == part else contents)
                    with self.assertRaises((AssertionError, KeyError)):
                        verify.inspect(candidate, 'Archway')


if __name__ == '__main__':
    unittest.main()
