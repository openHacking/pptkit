import importlib.util
import io
import unittest
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('native_mp4', Path(__file__).with_name('verify-native-mp4.py'))
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


class NativeMp4Test(unittest.TestCase):
    def test_sample_and_negative_controls(self):
        source = (ROOT / 'tests/fixtures/h264-mp4-powerpoint.pptx').read_bytes()
        self.assertEqual(oracle.verify(source)['objectId'], '4')
        for part, old, new in [
            ('ppt/slides/slide1.xml', b'<p:spTgt spid="4"', b'<p:spTgt spid="3"'),
            ('ppt/slides/_rels/slide1.xml.rels', b'../media/media1.mp4', b'../media/missing.mp4'),
            ('ppt/media/media1.mp4', b'ftyp', b'xxxx'),
            ('ppt/media/image1.png', b'PNG', b'BAD'),
            ('[Content_Types].xml', b'video/mp4', b'video/bad'),
        ]:
            with self.subTest(part=part, old=old), self.assertRaises((AssertionError, KeyError, AttributeError)):
                oracle.verify(mutate(source, part, old, new))


if __name__ == '__main__':
    unittest.main()
