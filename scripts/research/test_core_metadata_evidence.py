"""Adversarial checks for the retained metadata evidence verifier."""
import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
MANIFEST = ROOT / 'docs/testing/core-metadata-libreoffice-evidence.json'
VERIFIER = ROOT / 'scripts/research/verify-core-metadata-libreoffice.py'

@unittest.skipUnless(MANIFEST.exists(), 'Local evidence manifest unavailable')
class EvidenceVerificationTest(unittest.TestCase):
    def setUp(self):
        self.evidence = json.loads(MANIFEST.read_text())
        if not (ROOT / self.evidence['records'][0]['path']).exists():
            self.skipTest('Retained local artifacts unavailable')

    def run_verifier(self, evidence, optimized=False):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'evidence.json'
            path.write_text(json.dumps(evidence))
            return subprocess.run([sys.executable, *(['-O'] if optimized else []), str(VERIFIER), str(path)], capture_output=True, text=True)

    def test_original(self):
        result = self.run_verifier(self.evidence)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_reject_tampering_even_in_optimized_mode(self):
        mutations = []
        item = copy.deepcopy(self.evidence); item['records'][0]['sha256'] = '0' * 64; mutations.append(item)
        item = copy.deepcopy(self.evidence); item['records'][0]['inventory']['properties'][0]['value'] = 'tampered'; mutations.append(item)
        item = copy.deepcopy(self.evidence); item['records'][0]['path'] = '../outside.pptx'; mutations.append(item)
        item = copy.deepcopy(self.evidence); item['records'][0]['round'] = 3; mutations.append(item)
        item = copy.deepcopy(self.evidence); item['sourceChangedParts'] = []; mutations.append(item)
        item = copy.deepcopy(self.evidence); item['oracleSha256'] = '0' * 64; mutations.append(item)
        for item in mutations:
            with self.subTest(item=item):
                self.assertNotEqual(self.run_verifier(item, optimized=True).returncode, 0)

if __name__ == '__main__':
    unittest.main()
