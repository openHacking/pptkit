import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / 'docs/testing/drawing-properties-libreoffice-evidence.json'
VERIFIER = ROOT / 'scripts/research/verify-drawing-properties-libreoffice.py'


class RetainedEvidenceVerifierTests(unittest.TestCase):
    def test_rejects_tampered_records_even_under_optimized_python(self):
        if not EVIDENCE.exists():
            self.skipTest('Retained producer experiment has not been recorded')
        evidence = json.loads(EVIDENCE.read_text())
        if not all((ROOT / record['path']).exists() for record in evidence['rounds']):
            self.skipTest('Retained producer files are not present on this machine')
        cases = []
        bad_hash = copy.deepcopy(evidence); bad_hash['rounds'][1]['sha256'] = '0' * 64; cases.append(bad_hash)
        bad_inventory = copy.deepcopy(evidence); bad_inventory['rounds'][1]['inventory']['firstSlideNumber'] = '999'; cases.append(bad_inventory)
        bad_path = copy.deepcopy(evidence); bad_path['rounds'][1]['path'] = '../../etc/passwd'; cases.append(bad_path)
        bad_round = copy.deepcopy(evidence); bad_round['rounds'][1]['round'] = 2; cases.append(bad_round)
        with tempfile.TemporaryDirectory() as directory:
            for index, item in enumerate(cases):
                manifest = Path(directory) / f'{index}.json'; manifest.write_text(json.dumps(item))
                result = subprocess.run([sys.executable, '-O', str(VERIFIER), str(manifest)], capture_output=True, text=True, timeout=15)
                self.assertNotEqual(result.returncode, 0, result.stdout)


if __name__ == '__main__':
    unittest.main()
