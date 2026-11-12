import importlib.util
import unittest
from pathlib import Path

import fitz

spec = importlib.util.spec_from_file_location('render_audit', Path(__file__).with_name('corpus-headless-render-audit.py'))
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)


class RenderAuditTest(unittest.TestCase):
    def test_equal_changed_and_page_bounds(self):
        a, b = fitz.open(), fitz.open()
        first, second = a.new_page(width=100, height=100), b.new_page(width=100, height=100)
        equal = audit.pixels(first, second)
        self.assertEqual(equal['changedPixels'], 0)
        self.assertIsNone(equal['changedBoundsInclusive'])
        second.draw_rect(fitz.Rect(10, 10, 20, 20), color=(0, 0, 0), fill=(0, 0, 0))
        changed = audit.pixels(first, second)
        self.assertGreater(changed['changedPixels'], 0)
        self.assertGreater(changed['maxChannelDelta'], 0)
        self.assertIsNotNone(changed['changedBoundsInclusive'])
        c = fitz.open()
        third = c.new_page(width=101, height=100)
        self.assertEqual(audit.pixels(first, third)['status'], 'different-page-rectangles')
        a.close(); b.close(); c.close()


if __name__ == '__main__':
    unittest.main()
