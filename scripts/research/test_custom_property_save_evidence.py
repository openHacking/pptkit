"""Adversarial checks for paired custom-property evidence."""
import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
MANIFEST = ROOT / 'docs/testing/custom-property-save-libreoffice-evidence.json'
VERIFIER = ROOT / 'scripts/research/verify-custom-property-save-libreoffice.py'
@unittest.skipUnless(MANIFEST.exists(), 'local evidence unavailable')
class VerifyEvidence(unittest.TestCase):
    def setUp(self):
        self.data = json.loads(MANIFEST.read_text())
        if not (ROOT / self.data['records'][0]['path']).exists(): self.skipTest('local artifacts unavailable')
    def verify(self, data):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'manifest.json'; p.write_text(json.dumps(data))
            return subprocess.run([sys.executable, '-O', str(VERIFIER), str(p)], capture_output=True, text=True)
    def test_valid(self):
        result = self.verify(self.data); self.assertEqual(result.returncode, 0, result.stderr)
    def test_tampering_rejected(self):
        cases = []
        x = copy.deepcopy(self.data); x['records'][0]['sha256'] = '0' * 64; cases.append(x)
        x = copy.deepcopy(self.data); x['records'][0]['properties'][0]['value'] = 'wrong'; cases.append(x)
        x = copy.deepcopy(self.data); x['records'][0]['path'] = '../outside.pptx'; cases.append(x)
        x = copy.deepcopy(self.data); x['records'][0]['round'] = 3; cases.append(x)
        x = copy.deepcopy(self.data); x['sourceChangedParts'] = []; cases.append(x)
        x = copy.deepcopy(self.data); x['runnerSha256'] = '0' * 64; cases.append(x)
        for item in cases:
            with self.subTest(item=item): self.assertNotEqual(self.verify(item).returncode, 0)
if __name__ == '__main__': unittest.main()
