import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path

import fitz

spec = importlib.util.spec_from_file_location('visibility_oracle', Path(__file__).with_name('slide-visibility-oracle.py'))
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)


class VisibilityOracleTests(unittest.TestCase):
    def test_order_is_relationship_order_and_absent_is_distinct_from_explicit(self):
        p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        pr = 'http://schemas.openxmlformats.org/package/2006/relationships'
        with tempfile.TemporaryDirectory() as directory:
            file = Path(directory) / 'fixture.pptx'
            with zipfile.ZipFile(file, 'w') as archive:
                archive.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{p}" xmlns:r="{r}"><p:sldIdLst><p:sldId id="100" r:id="second"/><p:sldId id="101" r:id="first"/></p:sldIdLst></p:presentation>')
                archive.writestr('ppt/_rels/presentation.xml.rels', f'<Relationships xmlns="{pr}"><Relationship Id="first" Type="{r}/slide" Target="slides/slide1.xml"/><Relationship Id="second" Type="{r}/slide" Target="/ppt/slides/slide2.xml"/></Relationships>')
                archive.writestr('ppt/slides/slide1.xml', f'<p:sld xmlns:p="{p}" xmlns:a="{a}"><a:t>First</a:t></p:sld>')
                archive.writestr('ppt/slides/slide2.xml', f'<p:sld xmlns:p="{p}" xmlns:a="{a}" show="0"><a:t>Second</a:t></p:sld>')
            slides = oracle.inspect_pptx(file)['slides']
            self.assertEqual([s['text'] for s in slides], ['Second', 'First'])
            self.assertEqual([s['show'] for s in slides], ['0', None])
            self.assertEqual([s['hidden'] for s in slides], [True, False])
            pdf = Path(directory) / 'sample.pdf'
            doc = fitz.open(); page = doc.new_page(); page.insert_text((30, 30), 'First page'); doc.save(pdf); doc.close()
            report = oracle.inspect_pdf(pdf); self.assertEqual(report['pageCount'], 1); self.assertIn('First page', report['pages'][0]['text'])

    def test_duplicate_relationship_id_is_not_silently_selected(self):
        p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        pr = 'http://schemas.openxmlformats.org/package/2006/relationships'
        with tempfile.TemporaryDirectory() as directory:
            file = Path(directory) / 'ambiguous.pptx'
            with zipfile.ZipFile(file, 'w') as archive:
                archive.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{p}" xmlns:r="{r}"><p:sldIdLst><p:sldId id="100" r:id="slide"/></p:sldIdLst></p:presentation>')
                archive.writestr('ppt/_rels/presentation.xml.rels', f'<Relationships xmlns="{pr}"><Relationship Id="slide" Type="{r}/slide" Target="a.xml"/><Relationship Id="slide" Type="{r}/slide" Target="b.xml"/></Relationships>')
            with self.assertRaises(ValueError):
                oracle.inspect_pptx(file)

    def test_rejects_external_and_wrong_type_slide_relationships(self):
        p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        pr = 'http://schemas.openxmlformats.org/package/2006/relationships'
        with tempfile.TemporaryDirectory() as directory:
            for index, attrs in enumerate([f'Type="{r}/slide" TargetMode="External"', f'Type="{r}/notesSlide"']):
                file = Path(directory) / f'{index}.pptx'
                with zipfile.ZipFile(file, 'w') as archive:
                    archive.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{p}" xmlns:r="{r}"><p:sldIdLst><p:sldId id="100" r:id="slide"/></p:sldIdLst></p:presentation>')
                    archive.writestr('ppt/_rels/presentation.xml.rels', f'<Relationships xmlns="{pr}"><Relationship Id="slide" {attrs} Target="slides/slide1.xml"/></Relationships>')
                with self.assertRaises(ValueError):
                    oracle.inspect_pptx(file)


if __name__ == '__main__':
    unittest.main()
