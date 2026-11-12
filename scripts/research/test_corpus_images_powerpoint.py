import runpy
import tempfile
import unittest
import zipfile
from pathlib import Path

module = runpy.run_path(str(Path(__file__).with_name("verify-corpus-images-powerpoint.py")))


class CorpusImagesTests(unittest.TestCase):
    def test_retained_fixture(self):
        module["verify"]()

    def test_mutations_are_rejected(self):
        path = module["ROOT"] / "tests/fixtures/corpus-images-powerpoint-interactive.pptx"
        with zipfile.ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        changes = [("ppt/slides/slide1.xml", b'18212', b'18213'),
                   ("ppt/slides/slide1.xml", b'924338', b'924339'),
                   ("ppt/slides/slide1.xml", b'image-cover', b'image-other'),
                   ("ppt/slides/slide1.xml", b'l="10000" t="15000"', b'l="10001" t="15000"'),
                   ("ppt/media/image1.png", b'PNG', b'PNH')]
        media_hash = "45ca74602a2f7790a8eb2fad662cc73ee8ab42e78d5abfd0c66aa4d44725c9c6"
        for part, before, after in changes:
            with self.subTest(part=part, before=before), tempfile.TemporaryDirectory() as directory:
                self.assertIn(before, parts[part])
                candidate = Path(directory) / "candidate.pptx"
                with zipfile.ZipFile(candidate, "w") as archive:
                    for name, content in parts.items():
                        archive.writestr(name, content.replace(before, after, 1) if name == part else content)
                with self.assertRaises(ValueError):
                    module["inspect"](candidate, True, media_hash)


if __name__ == "__main__":
    unittest.main()
