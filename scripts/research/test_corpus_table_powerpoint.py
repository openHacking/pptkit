import copy
import runpy
import unittest
from pathlib import Path


module = runpy.run_path(str(Path(__file__).with_name("verify-corpus-table-powerpoint.py")))


class CorpusTablePowerPointTests(unittest.TestCase):
    def test_retained_fixture(self):
        module["verify"]()
        module["verify"](fresh=True)

    def test_changed_text_and_each_merge_are_rejected(self):
        fixture = module["ROOT"] / "tests/fixtures/corpus-table-powerpoint-interactive.pptx"
        cells = module["inventory"](fixture)
        changed = copy.deepcopy(cells)
        changed[1][0] = (changed[1][0][0], "wrong")
        with self.assertRaises(ValueError):
            module["check_table"](changed, "Edited rows")
        for row, column, key in [(0, 0, "gridSpan"), (0, 1, "hMerge"), (0, 2, "hMerge"),
                                 (1, 0, "rowSpan"), (2, 0, "vMerge")]:
            changed = copy.deepcopy(cells)
            del changed[row][column][0][key]
            with self.subTest(key=key, column=column), self.assertRaises(ValueError):
                module["check_table"](changed, "Edited rows")


if __name__ == "__main__":
    unittest.main()
