import importlib.util
from io import BytesIO
from pathlib import Path
import unittest
from zipfile import ZipFile

spec = importlib.util.spec_from_file_location("elbow_move", Path(__file__).with_name("verify-native-elbow-move.py"))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


def change_slide(data, old, new):
    result = BytesIO()
    with ZipFile(BytesIO(data)) as source, ZipFile(result, "w") as target:
        for entry in source.infolist():
            content = source.read(entry.filename)
            if entry.filename == "ppt/slides/slide1.xml":
                if old not in content:
                    raise AssertionError("mutation target absent")
                content = content.replace(old, new, 1)
            target.writestr(entry, content)
    return result.getvalue()


class NativeElbowMove(unittest.TestCase):
    def test_observation_and_negative_controls(self):
        before, after = [(module.base.ROOT / "tests/fixtures" / name).read_bytes() for name in module.HASHES]
        module.check(before, after)
        for old, new in [
            (b'bentConnector3', b'bentConnector2'),
            (b'val 50000', b'val 25000'),
            (b'<a:endCxn id="4" idx="1"', b'<a:endCxn id="4" idx="0"'),
            (b'cx="4207565"', b'cx="4207564"'),
        ]:
            with self.assertRaises(ValueError):
                module.check(before, change_slide(after, old, new))


if __name__ == "__main__":
    unittest.main()
