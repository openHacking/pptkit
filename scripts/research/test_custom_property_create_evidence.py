"""Adversarial verification of recorded custom-property creation evidence."""
import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
MANIFEST = ROOT / 'docs/testing/custom-property-create-libreoffice-evidence.json'
VERIFIER = ROOT / 'scripts/research/verify-custom-property-create-libreoffice.py'
@unittest.skipUnless(MANIFEST.exists(), 'local evidence unavailable')
class VerifyEvidence(unittest.TestCase):
    def setUp(self):
        self.data = json.loads(MANIFEST.read_text())
        if not (ROOT / self.data['records'][0]['path']).exists(): self.skipTest('local artifacts unavailable')
    def verify(self, data):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'manifest.json'; path.write_text(json.dumps(data))
            return subprocess.run([sys.executable, '-O', str(VERIFIER), str(path)], capture_output=True, text=True)
    def test_valid(self):
        result = self.verify(self.data); self.assertEqual(result.returncode, 0, result.stderr)
    def test_tampering_rejected(self):
        cases = []
        for mutation in ('hash', 'value', 'path', 'round', 'source', 'runner', 'log'):
            item = copy.deepcopy(self.data)
            if mutation == 'hash': item['records'][0]['sha256'] = '0' * 64
            if mutation == 'value': item['records'][1]['properties'][0]['value'] = 'wrong'
            if mutation == 'path': item['records'][0]['path'] = '../outside.pptx'
            if mutation == 'round': item['records'][0]['round'] = 3
            if mutation == 'source': item['sourceChangedParts'] = []
            if mutation == 'runner': item['runnerSha256'] = '0' * 64
            if mutation == 'log': item['logs'][0]['round'] = 3
            cases.append(item)
        for item in cases:
            with self.subTest(item=item): self.assertNotEqual(self.verify(item).returncode, 0)
if __name__ == '__main__': unittest.main()
