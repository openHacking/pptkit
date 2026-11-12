import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path

spec = importlib.util.spec_from_file_location('verifier', Path(__file__).with_name('verify-corpus-charts-libreoffice.py'))
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)
SAVED = audit.ROOT / 'tests/fixtures/corpus-charts-libreoffice-interactive.pptx'


class LibreOfficeChartEditEvidenceTest(unittest.TestCase):
    def test_saved_fixture(self):
        audit.verify()

    def test_rejects_wrong_edit_control_reference_or_workbook(self):
        for name, replace, message in [
            ('ppt/charts/chart1.xml', lambda data: data.replace(b'<c:v>1.5</c:v>', b'<c:v>2</c:v>'), 'chart values in bar'),
            ('ppt/charts/chart2.xml', lambda data: data.replace(b'<c:v>4</c:v>', b'<c:v>9</c:v>'), 'chart values in line'),
            ('ppt/slides/_rels/slide1.xml.rels',
             lambda data: data.replace(b'../charts/chart1.xml', b'../charts/chart2.xml'), 'chart identity'),
            ('ppt/embeddings/decoy.xlsx', lambda data: b'not a retained workbook', 'embedded worksheet'),
        ]:
            with self.subTest(name=name), tempfile.TemporaryDirectory() as scratch:
                path = Path(scratch) / 'mutation.pptx'
                with zipfile.ZipFile(SAVED) as source, zipfile.ZipFile(path, 'w') as target:
                    for item in source.namelist():
                        data = source.read(item)
                        target.writestr(item, replace(data) if item == name else data)
                    if name not in source.namelist():
                        target.writestr(name, replace(b''))
                with self.assertRaisesRegex(ValueError, message):
                    audit.check_saved(path)


if __name__ == '__main__':
    unittest.main()
