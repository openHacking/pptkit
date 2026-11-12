import runpy
import tempfile
import unittest
import zipfile
from pathlib import Path

module = runpy.run_path(str(Path(__file__).with_name("verify-corpus-theme-powerpoint.py")))


class CorpusThemeTests(unittest.TestCase):
    def test_retained_fixture(self):
        module["verify"]()

    def test_mutations_are_rejected(self):
        path = module["ROOT"] / "tests/fixtures/corpus-theme-powerpoint-interactive.pptx"
        with zipfile.ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        changes = [
            ("ppt/slides/slide1.xml", b"Edited", b"Wrong"),
            ("ppt/slides/slide2.xml", b'show="0"', b'show="1"'),
            ("ppt/slides/_rels/slide1.xml.rels", b"slideLayout2.xml", b"slideLayout1.xml"),
            ("ppt/slideLayouts/_rels/slideLayout2.xml.rels", b"slideMaster1.xml", b"slideMaster2.xml"),
            ("ppt/slideMasters/_rels/slideMaster1.xml.rels", b"theme1.xml", b"theme2.xml"),
        ]
        for part, before, after in changes:
            with self.subTest(part=part), tempfile.TemporaryDirectory() as directory:
                self.assertIn(before, parts[part])
                candidate = Path(directory) / "candidate.pptx"
                with zipfile.ZipFile(candidate, "w") as archive:
                    for name, content in parts.items():
                        archive.writestr(name, content.replace(before, after, 1) if name == part else content)
                with self.assertRaises(ValueError):
                    module["inspect"](candidate, "Edited title")


if __name__ == "__main__":
    unittest.main()
