import runpy
import tempfile
import unittest
import zipfile
from pathlib import Path

module = runpy.run_path(str(Path(__file__).with_name("verify-corpus-text-powerpoint.py")))


class CorpusTextPowerPointTests(unittest.TestCase):
    def test_retained_fixture(self):
        module["verify"]()

    def test_corrupted_text_list_hidden_notes_and_relationship_are_rejected(self):
        path = module["ROOT"] / "tests/fixtures/corpus-text-powerpoint-interactive.pptx"
        with zipfile.ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        changes = [
            ("ppt/slides/slide1.xml", b"Revised", b"Wrong"),
            ("ppt/slides/slide1.xml", b'startAt="3"', b'startAt="4"'),
            ("ppt/slides/slide1.xml", b"ppaction://hlinksldjump", b"ppaction://noaction"),
            ("ppt/slides/slide2.xml", b'show="0"', b'show="1"'),
            ("ppt/notesSlides/notesSlide2.xml", b"reviewer", b"other"),
            ("ppt/slides/_rels/slide1.xml.rels", b'Target="slide2.xml"', b'Target="slide1.xml"'),
        ]
        for name, before, after in changes:
            with self.subTest(name=name, change=before), tempfile.TemporaryDirectory() as directory:
                self.assertIn(before, parts[name])
                candidate = Path(directory) / "candidate.pptx"
                with zipfile.ZipFile(candidate, "w") as archive:
                    for part, content in parts.items():
                        archive.writestr(part, content.replace(before, after, 1) if part == name else content)
                with self.assertRaises(ValueError):
                    module["inspect"](candidate, True)


if __name__ == "__main__":
    unittest.main()
