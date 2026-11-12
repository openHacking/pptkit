import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / 'docs/testing/slide-visibility-libreoffice-evidence.json'
VERIFIER = ROOT / 'scripts/research/verify-slide-visibility-libreoffice.py'


class VisibilityEvidenceVerifierTests(unittest.TestCase):
    def test_rejects_tampering_even_with_optimized_python(self):
        if not EVIDENCE.exists():
            self.skipTest('No local retained producer record')
        evidence = json.loads(EVIDENCE.read_text())
        if not all((ROOT / record['pptx']['path']).exists() and (ROOT / record['pdf']['path']).exists() for record in evidence['records']):
            self.skipTest('Retained producer files are absent')
        cases = []
        wrong_hash = copy.deepcopy(evidence); wrong_hash['records'][1]['pdf']['sha256'] = '0' * 64; cases.append(wrong_hash)
        wrong_inventory = copy.deepcopy(evidence); wrong_inventory['records'][1]['inventory']['pptx']['slides'][1]['hidden'] = False; cases.append(wrong_inventory)
        wrong_path = copy.deepcopy(evidence); wrong_path['records'][1]['pptx']['path'] = '../../etc/passwd'; cases.append(wrong_path)
        wrong_order = copy.deepcopy(evidence); wrong_order['records'][1]['round'] = 2; cases.append(wrong_order)
        wrong_source = copy.deepcopy(evidence); wrong_source['sourceChangedParts'] = []; cases.append(wrong_source)
        with tempfile.TemporaryDirectory() as directory:
            for index, data in enumerate(cases):
                file = Path(directory) / f'{index}.json'; file.write_text(json.dumps(data))
                result = subprocess.run([sys.executable, '-O', str(VERIFIER), str(file)], capture_output=True, text=True, timeout=15)
                self.assertNotEqual(result.returncode, 0, result.stdout)


if __name__ == '__main__':
    unittest.main()
