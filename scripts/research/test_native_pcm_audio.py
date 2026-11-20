import importlib.util
import io
import unittest
from pathlib import Path
from zipfile import ZipFile

path = Path(__file__).with_name("verify-native-pcm-audio.py")
spec = importlib.util.spec_from_file_location("native_pcm_audio", path)
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)


def mutate(source, target, old, new):
    output = io.BytesIO()
    with ZipFile(io.BytesIO(source)) as src, ZipFile(output, "w") as dst:
        for entry in src.infolist():
            data = src.read(entry.filename)
            if entry.filename == target:
                assert old in data
                data = data.replace(old, new, 1)
            dst.writestr(entry, data)
    return output.getvalue()


class NativeAudioOracleTest(unittest.TestCase):
    def test_mutations_are_detected(self):
        source = oracle.SOURCE.read_bytes()
        self.assertEqual(oracle.inspect(source)["frames"], 8000)
        cases = [
            ("ppt/slides/_rels/slide1.xml.rels", b"../media/media1.wav", b"../media/other.wav"),
            ("ppt/slides/_rels/slide1.xml.rels", b"relationships/audio", b"relationships/video"),
            ("ppt/slides/slide1.xml", b"playFrom(0.0)", b"playFrom(1.0)"),
            ("ppt/slides/slide1.xml", b'<p:spTgt spid="4"', b'<p:spTgt spid="5"'),
            ("[Content_Types].xml", b"audio/x-wav", b"audio/mpeg"),
            ("ppt/media/media1.wav", b"RIFF", b"RIFX"),
        ]
        for target, old, new in cases:
            with self.subTest(target=target, old=old):
                with self.assertRaises((AssertionError, KeyError, ValueError)):
                    oracle.inspect(mutate(source, target, old, new))


if __name__ == "__main__":
    unittest.main()
