import importlib.util
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location('reorder_oracle', Path(__file__).with_name('reorder-libreoffice-oracle.py'))
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)


def node(name, children=()):
    return {'name': name, 'children': list(children)}


class ReorderOracleTest(unittest.TestCase):
    def test_source_order_detects_top_and_both_group_levels(self):
        def sample(top, outer, inner):
            group = node('Outer', [node(outer[0]), node(outer[1]), node('Inner', [node(inner[0]), node(inner[1])])])
            return {'objects': [node(top[0]), node(top[1]), group, node('Caption')]}
        control = sample(['A', 'B'], ['C', 'D'], ['E', 'F'])
        changed = sample(['B', 'A'], ['D', 'C'], ['F', 'E'])
        oracle.check_source([control, changed])
        for bad in (sample(['A', 'B'], ['D', 'C'], ['F', 'E']), sample(['B', 'A'], ['C', 'D'], ['F', 'E']), sample(['B', 'A'], ['D', 'C'], ['E', 'F'])):
            with self.assertRaises(AssertionError):
                oracle.check_source([control, bad])

    def test_raster_diff_and_size_mismatch(self):
        image = {'width': 2, 'height': 1, 'components': 3, 'samples': b'\x00\x00\x00\xff\xff\xff'}
        changed = dict(image, samples=b'\x00\x00\x00\xff\xff\x00')
        self.assertEqual(oracle.difference(image, image)['changedPixels'], 0)
        self.assertEqual(oracle.difference(image, changed)['changedPixels'], 1)
        self.assertEqual(oracle.difference(image, dict(image, width=1))['status'], 'size-mismatch')

    def test_named_geometry_does_not_silently_merge_duplicate_names(self):
        shape = {'name': 'A', 'transform': {'ext': {'cx': '100', 'cy': '200'}}, 'children': []}
        group = {'name': 'Group', 'transform': None, 'children': [shape]}
        self.assertEqual(oracle.named_transforms([group])['A']['ext']['cx'], '100')
        with self.assertRaisesRegex(ValueError, 'Duplicate name'):
            oracle.named_transforms([group, shape])


if __name__ == '__main__':
    unittest.main()
