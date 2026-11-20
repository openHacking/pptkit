import importlib.util
import io
import unittest
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('multi_video', Path(__file__).with_name('verify-multi-video.py'))
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)


def mutate(source, part, old, new):
    output = io.BytesIO()
    with ZipFile(io.BytesIO(source)) as src, ZipFile(output, 'w') as dst:
        for entry in src.infolist():
            data = src.read(entry.filename)
            if entry.filename == part:
                assert old in data
                data = data.replace(old, new, 1)
            dst.writestr(entry, data)
    return output.getvalue()


class MultiVideoTest(unittest.TestCase):
    def test_native_and_generated_and_negative_controls(self):
        for name in ['h264-mp4-powerpoint-duplicate.pptx', 'h264-mp4-pptkit-multi.pptx']:
            result = oracle.verify((ROOT / 'tests/fixtures' / name).read_bytes())
            self.assertEqual(len(result['objects']), 2)
        source = (ROOT / 'tests/fixtures/h264-mp4-pptkit-multi.pptx').read_bytes()
        for part, old, new in [
            ('ppt/slides/slide1.xml', b'<p:spTgt spid="3"', b'<p:spTgt spid="2"'),
            ('ppt/slides/slide1.xml', b'<p:cTn id="22"', b'<p:cTn id="21"'),
            ('ppt/slides/_rels/slide1.xml.rels', b'../media/movie.mp4', b'../media/missing.mp4'),
            ('ppt/media/movie.mp4', b'ftyp', b'xxxx'),
        ]:
            with self.subTest(part=part, old=old), self.assertRaises((AssertionError, KeyError, AttributeError)):
                oracle.verify(mutate(source, part, old, new))


if __name__ == '__main__':
    unittest.main()
