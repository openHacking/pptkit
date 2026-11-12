"""Owned fixtures for the optional Python native-artifact extraction oracle."""
import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path

spec = importlib.util.spec_from_file_location("geometry_extractor", Path(__file__).with_name("extract-native-diagram-geometry.py"))
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)

SLIDE = '''<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr name="Owned diagram"/></p:nvGraphicFramePr><p:xfrm><a:off x="12700" y="25400"/><a:ext cx="1270000" cy="635000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphic></p:graphicFrame></p:sld>'''


def document(body):
    return '''<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">''' + body + '</office:document>'


def page(body):
    return document('<draw:page>' + body + '</draw:page>')


class GeometryExtractionTests(unittest.TestCase):
    def extract(self, native, source=SLIDE):
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            source_file, native_file = base / 'source.pptx', base / 'native.fodp'
            with zipfile.ZipFile(source_file, 'w') as archive:
                archive.writestr('ppt/slides/slide1.xml', source)
            native_file.write_text(native, encoding='utf-8')
            return oracle.extract(source_file, native_file)

    def test_frame_units_and_direct_shape_text(self):
        actual = self.extract(page('''<draw:g draw:name="Owned diagram"><draw:custom-shape draw:name="15" svg:x="1in" svg:y="2pt" svg:width="2.54cm" svg:height="25.4mm"><text:p>A<text:span>B</text:span>C<text:s text:c="3"/>D<text:tab/>E<text:line-break/>中文 😀 &amp;</text:p><draw:enhanced-geometry draw:type="ooxml-ellipse" draw:text-areas="?f1 ?f2"/></draw:custom-shape></draw:g>'''))
        self.assertEqual(actual['snapshot']['frame'], dict(x=1, y=2, width=100, height=50))
        shape = actual['snapshot']['shapes'][0]
        self.assertEqual(shape['paragraphs'], ['ABC   D\tE\n中文 😀 &'])
        self.assertEqual(shape['preset'], 'ooxml-ellipse')
        self.assertEqual(shape['textAreas'], '?f1 ?f2')
        self.assertEqual(actual['oracle'][0]['boundsPoints'], dict(x=72, y=2, width=72, height=72))

    def test_frame_text_box_and_empty_paragraph_are_not_lost(self):
        actual = self.extract(page('''<draw:g draw:name="Owned diagram"><draw:frame draw:name="detail"><draw:text-box><text:p>Detail<text:line-break/>line</text:p><text:p/></draw:text-box></draw:frame></draw:g>'''))
        self.assertEqual(actual['snapshot']['shapes'][0]['paragraphs'], ['Detail\nline', ''])
        self.assertEqual(actual['oracle'][0]['status'], 'incomplete')

    def test_nested_native_lists_preserve_paragraph_order_without_copying_child_drawing_text(self):
        actual = self.extract(page('''<draw:g draw:name="Owned diagram"><draw:custom-shape draw:name="detail"><text:list><text:list-header><text:p>Header</text:p></text:list-header><text:list-item><text:p>First</text:p><text:list><text:list-item><text:p>Nested</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>Last</text:p></text:list-item></text:list><draw:custom-shape><text:p>Other shape</text:p></draw:custom-shape></draw:custom-shape></draw:g>'''))
        self.assertEqual(actual['snapshot']['shapes'][0]['paragraphs'], ['Header', 'First', 'Nested', 'Last'])

    def test_all_ancestor_transforms_and_repeated_names_are_retained(self):
        actual = self.extract(page('''<draw:g draw:transform="scale (2 2)"><draw:g draw:name="Owned diagram" draw:transform="translate (1cm 2cm)"><draw:custom-shape draw:name="18" draw:transform="rotate (1)"/><draw:custom-shape draw:name="18"/></draw:g></draw:g>'''))
        a, b = actual['snapshot']['shapes']
        self.assertNotEqual(a['ordinal'], b['ordinal'])
        self.assertEqual(a['name'], b['name'])
        self.assertEqual(a['ancestorTransforms'], ['scale (2 2)', 'translate (1cm 2cm)'])
        self.assertEqual(a['transform'], 'rotate (1)')
        self.assertTrue(all(r['status'] == 'transform-source-only' for r in actual['oracle']))

    def test_missing_ambiguous_groups_and_master_artwork_do_not_become_empty_success(self):
        missing = self.extract(document('<draw:g draw:name="Owned diagram"/><draw:page><draw:g draw:name="different"/></draw:page>'))
        self.assertEqual(missing['groupAssociation']['status'], 'missing')
        self.assertEqual(missing['snapshot']['shapes'], [])
        ambiguous = self.extract(page('<draw:g draw:name="Owned diagram"/><draw:g draw:name="Owned diagram"/>'))
        self.assertEqual(ambiguous['groupAssociation']['status'], 'ambiguous')
        self.assertEqual(ambiguous['snapshot']['shapes'], [])
        with self.assertRaises(ValueError):
            self.extract(document('<draw:page/><draw:page/>'))

    def test_dtd_resource_limits_and_invalid_whitespace_reject(self):
        with self.assertRaises(ValueError):
            oracle.xml_bytes(b'<!DOCTYPE root [<!ENTITY x "boom">]><root>&x;</root>')
        with self.assertRaises(ValueError):
            oracle.xml_bytes(b' ' * (oracle.MAX_BYTES + 1))
        for count in ['-1', '0', '10001', 'invalid']:
            with self.assertRaises(ValueError):
                self.extract(page('<draw:g draw:name="Owned diagram"><draw:custom-shape><text:p><text:s text:c="' + count + '"/></text:p></draw:custom-shape></draw:g>'))

    def test_decimal_oracle_and_unknown_units(self):
        self.assertEqual(oracle.points('6pc'), 72)
        self.assertEqual(oracle.points('-2.54cm'), -72)
        self.assertEqual(oracle.points('0'), 0)
        for value in ['3px', '1e3pt', 'NaNpt', 'Infinitypt', '1000000001pt', None]:
            self.assertIsNone(oracle.points(value))


if __name__ == '__main__':
    unittest.main()
