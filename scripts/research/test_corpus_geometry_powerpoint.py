import runpy
import tempfile
import unittest
import zipfile
from pathlib import Path

module = runpy.run_path(str(Path(__file__).with_name("verify-corpus-geometry-powerpoint.py")))


class CorpusGeometryTests(unittest.TestCase):
    def test_retained_fixture(self):
        module["verify"]()

    def test_mutations_are_rejected(self):
        path = module["ROOT"] / "tests/fixtures/corpus-geometry-powerpoint-interactive.pptx"
        with zipfile.ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        slide = "ppt/slides/slide1.xml"
        for before, after in [(b'rot="600000"', b'rot="0"'), (b'flipH="1"', b'flipH="0"'),
                              (b'1359451', b'1359452'), (b'nested-shape', b'other-shape'),
                              (b'Nested', b'Changed'), (b'name="anchored"', b'name="other"')]:
            with self.subTest(before=before), tempfile.TemporaryDirectory() as directory:
                self.assertIn(before, parts[slide])
                candidate = Path(directory) / "candidate.pptx"
                with zipfile.ZipFile(candidate, "w") as archive:
                    for name, content in parts.items():
                        archive.writestr(name, content.replace(before, after, 1) if name == slide else content)
                with self.assertRaises(ValueError):
                    module["inspect"](candidate, [1359451, 3177208])


if __name__ == "__main__":
    unittest.main()
