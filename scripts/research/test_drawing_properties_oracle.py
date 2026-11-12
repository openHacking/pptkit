import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path

spec = importlib.util.spec_from_file_location('drawing_properties_oracle', Path(__file__).with_name('drawing-properties-oracle.py'))
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)


class DrawingPropertyOracleTests(unittest.TestCase):
    def test_distinguishes_absent_false_true_and_preserves_text_and_order(self):
        p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        dec = 'http://schemas.microsoft.com/office/drawing/2017/decorative'
        uri = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}'
        def shape(name, value):
            extension = '' if value is None else f'<a:extLst><a:ext uri="{uri}"><d:decorative val="{value}"/></a:ext></a:extLst>'
            return f'<p:sp><p:nvSpPr><p:cNvPr id="2" name="{name}" descr="Alt &amp; text">{extension}</p:cNvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>{name}</a:t></a:r></a:p></p:txBody></p:sp>'
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'fixture.pptx'
            with zipfile.ZipFile(path, 'w') as archive:
                archive.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{p}" firstSlideNum="17"/>')
                for number in (2, 1):
                    shapes = shape('Missing', None) + shape('False', '0') + shape('True', '1')
                    archive.writestr(f'ppt/slides/slide{number}.xml', f'<p:sld xmlns:p="{p}" xmlns:a="{a}" xmlns:d="{dec}"><p:cSld><p:spTree>{shapes}</p:spTree></p:cSld></p:sld>')
            report = oracle.inspect(path)
            self.assertEqual(report['firstSlideNumber'], '17')
            self.assertEqual([page['part'] for page in report['pages']], ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml'])
            for page in report['pages']:
                self.assertEqual([x['decorativeValues'] for x in page['objects']], [[], ['0'], ['1']])
                self.assertEqual([x['text'] for x in page['objects']], ['Missing', 'False', 'True'])
                self.assertTrue(all(x['description'] == 'Alt & text' for x in page['objects']))


if __name__ == '__main__':
    unittest.main()
