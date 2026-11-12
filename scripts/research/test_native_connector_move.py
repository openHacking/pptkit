import copy
import importlib.util
from pathlib import Path
import unittest

path = Path(__file__).with_name("verify-native-connector-move.py")
spec = importlib.util.spec_from_file_location("connector_move", path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class ConnectorMoveEvidence(unittest.TestCase):
    def test_retained_evidence_and_corruptions(self):
        files = [module.ROOT / "tests/fixtures" / name for name in module.HASHES]
        before, after = (module.inspect(file.read_bytes()) for file in files)
        module.check(before, after)
        for mutate in (
            lambda x: x["link"]["connections"][1].update(id="999"),
            lambda x: x["link"]["connections"][1].update(idx="0"),
            lambda x: x["link"].update(box=before["link"]["box"]),
            lambda x: x["right"].update(box=before["right"]["box"]),
            lambda x: x.pop("left"),
        ):
            corrupt = copy.deepcopy(after)
            mutate(corrupt)
            with self.assertRaises(ValueError):
                module.check(before, corrupt)


if __name__ == "__main__":
    unittest.main()
