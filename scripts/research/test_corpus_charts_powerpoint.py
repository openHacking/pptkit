import io
import runpy
import tempfile
import unittest
import zipfile
from pathlib import Path

module = runpy.run_path(str(Path(__file__).with_name("verify-corpus-charts-powerpoint.py")))


class CorpusChartsTests(unittest.TestCase):
    def test_retained_fixture(self):
        module["verify"]()

    def test_mutations_are_rejected(self):
        path = module["ROOT"] / "tests/fixtures/corpus-charts-powerpoint-interactive.pptx"
        with zipfile.ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        workbook_name = "ppt/embeddings/Microsoft_Excel____.xlsx"
        with zipfile.ZipFile(io.BytesIO(parts[workbook_name])) as workbook:
            sheets = {name: workbook.read(name) for name in workbook.namelist()}
        sheet = "xl/worksheets/sheet1.xml"
        self.assertIn(b'<v>2</v>', sheets[sheet])
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, "w") as workbook:
            for name, content in sheets.items():
                workbook.writestr(name, content.replace(b'<v>2</v>', b'<v>9</v>', 1) if name == sheet else content)
        mutations = [("ppt/charts/chart1.xml", b"<a:t>Edited</a:t>", b"<a:t>Wrong</a:t>"),
                     ("ppt/charts/chart1.xml", b'<c:v>2</c:v>', b'<c:v>9</c:v>'),
                     ("ppt/charts/_rels/chart1.xml.rels", b"Microsoft_Excel____.xlsx", b"Microsoft_Excel____1.xlsx")]
        for part, before, after in mutations + [(workbook_name, parts[workbook_name], stream.getvalue())]:
            with self.subTest(part=part, before=before[:30]), tempfile.TemporaryDirectory() as directory:
                self.assertIn(before, parts[part])
                candidate = Path(directory) / "candidate.pptx"
                with zipfile.ZipFile(candidate, "w") as archive:
                    for name, content in parts.items():
                        archive.writestr(name, content.replace(before, after, 1) if name == part else content)
                with self.assertRaises(ValueError):
                    module["inspect"](candidate, True)


if __name__ == "__main__":
    unittest.main()
