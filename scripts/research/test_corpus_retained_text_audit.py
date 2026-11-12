import hashlib
import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path

import fitz

spec = importlib.util.spec_from_file_location('corpus_audit', Path(__file__).with_name('corpus-retained-text-audit.py'))
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)


def metadata(path):
    data = path.read_bytes()
    return {'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}


def deck(path, text='Visible', hidden=False):
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{audit.P[1:-1]}" xmlns:r="{audit.R[1:-1]}"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>')
        archive.writestr('ppt/_rels/presentation.xml.rels', f'<Relationships xmlns="{audit.REL[1:-1]}"><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>')
        archive.writestr('ppt/slides/slide1.xml', f'<p:sld xmlns:p="{audit.P[1:-1]}" xmlns:a="{audit.A[1:-1]}" show="{0 if hidden else 1}"><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:sld>')


class RetainedAuditTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.old_root = audit.ROOT
        audit.ROOT = Path(self.temp.name)
        for folder in ('.tmp/fixtures', '.tmp/compatibility/roundtrip', '.tmp/compatibility/reopened'):
            (audit.ROOT / folder).mkdir(parents=True)
        self.source = audit.ROOT / '.tmp/fixtures/x.pptx'
        self.saved = audit.ROOT / '.tmp/compatibility/roundtrip/x.pptx'
        self.pdf = audit.ROOT / '.tmp/compatibility/reopened/x.pdf'
        deck(self.source)
        deck(self.saved)
        self.make_pdf('Visible')

    def tearDown(self):
        audit.ROOT = self.old_root
        self.temp.cleanup()

    def make_pdf(self, text):
        pdf = fitz.open()
        page = pdf.new_page()
        page.insert_text((30, 40), text)
        pdf.save(self.pdf)
        pdf.close()

    def record(self):
        return {'fixture': 'x', 'input': metadata(self.source), 'saved': metadata(self.saved), 'reopenedPdf': metadata(self.pdf)}

    def test_identity_hashes_and_text_mismatch(self):
        record = self.record()
        result = audit.audit(record)
        self.assertEqual(result['textObservations'][0]['missingExactWithoutWhitespace'], [])
        self.assertTrue(result['slideComparisons'][0]['paragraphsExact'])
        self.saved.write_bytes(b'tampered')
        with self.assertRaisesRegex(ValueError, 'hash or length mismatch'):
            audit.audit(record)
        deck(self.saved, text='Different')
        result = audit.audit(self.record())
        self.assertFalse(result['slideComparisons'][0]['paragraphsExact'])
        self.assertEqual(result['textObservations'][0]['missingExactWithoutWhitespace'], ['Different'])

    def test_hidden_page_does_not_get_guessed_pdf_mapping(self):
        deck(self.saved, hidden=True)
        result = audit.audit(self.record())
        self.assertEqual(result['pageMapping'], 'unresolved-page-count')
        self.assertEqual(result['textObservations'], [])
        self.assertFalse(result['slideComparisons'][0]['hiddenEqual'])


if __name__ == '__main__':
    unittest.main()
