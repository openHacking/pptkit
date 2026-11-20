import importlib.util
import io
import unittest
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('multi_audio', Path(__file__).with_name('verify-multi-audio.py'))
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


class MultiAudioTest(unittest.TestCase):
    def test_oracle_rejects_broken_identity_and_relationships(self):
        source = (ROOT / 'tests/fixtures/pcm-wave-powerpoint-duplicate.pptx').read_bytes()
        self.assertEqual(oracle.verify(source)['objects'], ['2', '3'])
        for part, old, new in [
            ('ppt/slides/slide1.xml', b'<p:spTgt spid="3"', b'<p:spTgt spid="2"'),
            ('ppt/slides/slide1.xml', b'<p:cTn id="12"', b'<p:cTn id="11"'),
            ('ppt/slides/_rels/slide1.xml.rels', b'../media/media1.wav', b'../media/missing.wav'),
            ('ppt/media/media1.wav', b'RIFF', b'RIFX'),
        ]:
            with self.subTest(part=part, old=old), self.assertRaises((AssertionError, KeyError, AttributeError)):
                oracle.verify(mutate(source, part, old, new))


if __name__ == '__main__':
    unittest.main()
