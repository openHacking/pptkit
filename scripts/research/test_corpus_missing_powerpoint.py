import runpy
import tempfile
import unittest
import zipfile
from pathlib import Path

module = runpy.run_path(str(Path(__file__).with_name("verify-corpus-missing-powerpoint.py")))


class CorpusMissingTests(unittest.TestCase):
    def test_retained_fixture(self):
        module["verify"]()

    def test_injected_picture_media_or_relationship_is_rejected(self):
        path = module["ROOT"] / "tests/fixtures/corpus-missing-powerpoint-resave.pptx"
        with zipfile.ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        cases = [
            ("ppt/slides/slide1.xml", b"</p:spTree>", b"<p:pic/></p:spTree>"),
            ("ppt/slides/_rels/slide1.xml.rels", b"</Relationships>", b'<Relationship Id="x" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/x.png"/></Relationships>'),
        ]
        for part, before, after in cases:
            with self.subTest(part=part), tempfile.TemporaryDirectory() as directory:
                self.assertIn(before, parts[part])
                candidate = Path(directory) / "candidate.pptx"
                with zipfile.ZipFile(candidate, "w") as archive:
                    for name, content in parts.items():
                        archive.writestr(name, content.replace(before, after, 1) if name == part else content)
                with self.assertRaises(ValueError):
                    module["inspect"](candidate)
        with tempfile.TemporaryDirectory() as directory:
            candidate = Path(directory) / "candidate.pptx"
            with zipfile.ZipFile(candidate, "w") as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)
                archive.writestr("ppt/media/stray.png", b"stray")
            with self.assertRaises(ValueError):
                module["inspect"](candidate)


if __name__ == "__main__":
    unittest.main()
