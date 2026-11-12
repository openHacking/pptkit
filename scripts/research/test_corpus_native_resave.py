import importlib.util
import tempfile
import unittest
import warnings
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

spec = importlib.util.spec_from_file_location('native_resave', Path(__file__).with_name('corpus-native-resave.py'))
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)


class NativeResaveInventoryTest(unittest.TestCase):
    def setUp(self):
        scratch = tempfile.TemporaryDirectory()
        self.addCleanup(scratch.cleanup)
        self.path = Path(scratch.name) / 'case.pptx'

    def specimen(self):
        with zipfile.ZipFile(self.path, 'w') as archive:
            archive.writestr('ppt/presentation.xml',
                '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<p:sldIdLst><p:sldId r:id="link"/></p:sldIdLst></p:presentation>')
            archive.writestr('ppt/_rels/presentation.xml.rels',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="link" Target="slides/slide2.xml"/></Relationships>')
            archive.writestr('ppt/slides/slide2.xml',
                '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" show="0"/>')
            archive.writestr('ppt/charts/chart7.xml',
                '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                '<c:numCache><c:pt idx="2"><c:v>3.50</c:v></c:pt></c:numCache></c:chartSpace>')
            archive.writestr('ppt/embeddings/workbook.xlsx', b'workbook candidate')

    def test_reads_relationship_order_cache_indices_and_workbook_part_names(self):
        self.specimen()
        result = audit.inspect(self.path)
        self.assertEqual(result['slides'], [{'part': 'ppt/slides/slide2.xml', 'hidden': True, 'paragraphs': []}])
        self.assertEqual(result['chartNumericCaches'], {'ppt/charts/chart7.xml': [[['2', '3.50']]]})
        self.assertEqual(result['embeddedXlsxParts'], ['ppt/embeddings/workbook.xlsx'])

    def test_rejects_malformed_xml_even_when_unreferenced(self):
        self.specimen()
        with zipfile.ZipFile(self.path, 'a') as archive:
            archive.writestr('ppt/extra.xml', '<incomplete>')
        with self.assertRaises(ET.ParseError):
            audit.inspect(self.path)

    def test_rejects_crc_damage_and_duplicate_names(self):
        self.specimen()
        self.path.write_bytes(self.path.read_bytes().replace(b'workbook candidate', b'workbook corrupted'))
        with self.assertRaisesRegex(ValueError, 'CRC failure'):
            audit.inspect(self.path)
        self.specimen()
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', UserWarning)
            with zipfile.ZipFile(self.path, 'a') as archive:
                archive.writestr('ppt/charts/chart7.xml', '<decoy/>')
        with self.assertRaisesRegex(ValueError, 'Duplicate ZIP entry'):
            audit.inspect(self.path)


if __name__ == '__main__':
    unittest.main()
