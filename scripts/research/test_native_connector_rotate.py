import importlib.util
from pathlib import Path
import unittest

directory = Path(__file__).parent
spec = importlib.util.spec_from_file_location("rotate", directory / "verify-native-connector-rotate.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
helper_spec = importlib.util.spec_from_file_location("elbow_test", directory / "test_native_elbow_move.py")
helper = importlib.util.module_from_spec(helper_spec)
helper_spec.loader.exec_module(helper)


class NativeRotation(unittest.TestCase):
    def test_observation_and_negative_controls(self):
        before, after = [(module.base.ROOT / "tests/fixtures" / name).read_bytes() for name in module.HASHES]
        module.check(before, after)
        for old, new in [
            (b'rot="2727718"', b'rot="2727719"'),
            (b'flipV="1"', b'flipV="0"'),
            (b'<a:endCxn id="4" idx="1"', b'<a:endCxn id="4" idx="2"'),
            (b'cx="4094433"', b'cx="4094533"'),
        ]:
            with self.assertRaises(ValueError):
                module.check(before, helper.change_slide(after, old, new))


if __name__ == "__main__":
    unittest.main()
