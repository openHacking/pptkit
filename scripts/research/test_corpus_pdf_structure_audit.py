import importlib.util
import unittest
from pathlib import Path

import fitz

spec = importlib.util.spec_from_file_location('pdf_structure', Path(__file__).with_name('corpus-pdf-structure-audit.py'))
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)


class PdfStructureTest(unittest.TestCase):
    def test_text_geometry_path_and_count_changes(self):
        left, right = fitz.open(), fitz.open()
        a, b = left.new_page(width=100, height=100), right.new_page(width=100, height=100)
        a.insert_text((10, 20), 'Example')
        b.insert_text((11, 20), 'Example')
        a.draw_rect(fitz.Rect(10, 30, 20, 40))
        b.draw_rect(fitz.Rect(11, 30, 21, 40))
        result = audit.compare_pages(a, b)
        self.assertTrue(result['text']['extractedTextEqual'])
        self.assertEqual(result['text']['ordinalSpans'][0]['originMaxDelta'], 1)
        self.assertEqual(result['drawings']['ordinalPaths'][0]['rectMaxDelta'], 1)
        self.assertIn('items', result['drawings']['ordinalPaths'][0]['differentFields'])
        b.insert_text((5, 70), 'Extra')
        b.draw_rect(fitz.Rect(30, 30, 40, 40))
        result = audit.compare_pages(a, b)
        self.assertNotIn('ordinalSpans', result['text'])
        self.assertNotIn('ordinalPaths', result['drawings'])
        self.assertFalse(result['text']['extractedTextEqual'])
        left.close(); right.close()


if __name__ == '__main__':
    unittest.main()
