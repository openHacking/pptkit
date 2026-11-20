import importlib.util
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

spec = importlib.util.spec_from_file_location("mixed", Path(__file__).with_name("verify-mixed-media-comments.py"))
verifier = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verifier)


class MixedControls(unittest.TestCase):
    def test_corrupted_media_and_comment_references_fail(self):
        with ZipFile(verifier.DEFAULT) as source:
            parts = {item.filename: source.read(item) for item in source.infolist()}
        cases = [
            ("ppt/media/tone.wav", b"WAVEfmt ", b"FAILfmt "),
            ("ppt/comments/comment1.xml", b'authorId="0"', b'authorId="1"'),
            ("ppt/commentAuthors.xml", b'Fixture Author', b'Wrong Authorxx'),
            ("ppt/slides/slide1.xml", b'r:link="smAudio"', b'r:link="noAudio"'),
            ("ppt/slides/_rels/slide1.xml.rels", b'../media/tone.wav', b'../media/lost.wav'),
            ("ppt/slides/slide1.xml", b'<a:audioFile r:link="smAudio"/></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId4"', b'<a:audioFile r:link="smAudio"/></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="badId"'),
            ("[Content_Types].xml", b'ContentType="audio/wav"', b'ContentType="audio/mp3"'),
        ]
        with tempfile.TemporaryDirectory() as temp:
            for index, (part, before, after) in enumerate(cases):
                with self.subTest(index=index):
                    self.assertIn(before, parts[part])
                    target = Path(temp) / f"bad-{index}.pptx"
                    with ZipFile(target, "w") as archive:
                        for name, data in parts.items():
                            archive.writestr(name, data.replace(before, after, 1) if name == part else data)
                    with self.assertRaises((ValueError, KeyError)):
                        verifier.inspect(target)


if __name__ == "__main__":
    unittest.main()
