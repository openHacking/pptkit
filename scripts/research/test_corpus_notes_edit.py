import importlib.util
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

spec = importlib.util.spec_from_file_location('notes_edit', Path(__file__).with_name('verify-corpus-notes-edit.py'))
verify = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verify)


class NotesEditControls(unittest.TestCase):
    def test_corrupted_notes_graph_and_content_reject(self):
        with ZipFile(verify.OUTPUT) as z:
            original = {entry.filename: z.read(entry) for entry in z.infolist()}
        cases = [
            ('ppt/notesSlides/notesSlide2.xml', b'native', b'wrong'),
            ('ppt/slides/slide2.xml', b'show="0"', b'show="1"'),
            ('ppt/slides/_rels/slide2.xml.rels', b'notesSlide2.xml', b'notesSlide1.xml'),
            ('ppt/notesSlides/_rels/notesSlide2.xml.rels', b'slide2.xml', b'slide1.xml'),
            ('ppt/slides/_rels/slide1.xml.rels', b'slide2.xml', b'slide1.xml'),
        ]
        with tempfile.TemporaryDirectory() as temporary:
            for part, old, new in cases:
                with self.subTest(part=part):
                    self.assertIn(old, original[part])
                    path = Path(temporary) / 'invalid.pptx'
                    with ZipFile(path, 'w') as output:
                        for name, data in original.items():
                            output.writestr(name, data.replace(old, new, 1) if name == part else data)
                    with self.assertRaises((ValueError, KeyError)):
                        verify.inspect(path, 'Speaker & reviewer notes native edit')


if __name__ == '__main__':
    unittest.main()
