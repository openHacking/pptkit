import importlib.util
import io
from pathlib import Path
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('label_source', Path(__file__).with_name('verify-chart-label-source.py'))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

def replace(data, name, mutate):
    result = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(result, 'w', zipfile.ZIP_DEFLATED) as output:
        for entry in source.infolist():
            payload = source.read(entry.filename)
            output.writestr(entry, mutate(payload) if entry.filename == name else payload)
    return result.getvalue()

class LabelSourceTest(unittest.TestCase):
    def test_alias_and_escape_damage_is_detected(self):
        source = (ROOT / 'tests/fixtures/chart-label-source.pptx').read_bytes()
        self.assertEqual(module.inspect(source)['categoryCacheCopies'], 2)
        chart = 'ppt/charts/chart1.xml'
        book = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'
        cases = [
            replace(source, chart, lambda b: b.replace(b'literal _x005F_x0041_', b'wrong', 1)),
            replace(source, chart, lambda b: b.replace(b'Sheet1!$A$2:$A$4', b'Sheet1!$A$2:$A$3', 1)),
            replace(source, chart, lambda b: b.replace(b'literal _x005F_x0041_', b'literal _x0041_', 1)),
            replace(source, book, lambda b: replace(b, 'xl/worksheets/sheet1.xml', lambda x: x.replace(b' Q1 ', b'wrong', 1))),
        ]
        for damaged in cases:
            with self.subTest(size=len(damaged)), self.assertRaises(AssertionError):
                module.inspect(damaged)

if __name__ == '__main__':
    unittest.main()
