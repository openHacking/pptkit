#!/usr/bin/env python3
"""Record three headless LibreOfficeDev resaves of fresh boundary-size decks."""
import hashlib
import json
import platform
import subprocess
import tempfile
from pathlib import Path
from notes_size_inventory import inventory

ROOT = Path(__file__).resolve().parents[2]
SOFFICE = Path('/Users/openhacking/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice')
BASE = ROOT / '.tmp/notes-size-libreoffice'
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    BASE.mkdir(parents=True, exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix='run-', dir=BASE))
    records, logs = [], []
    for shape in ('portrait', 'landscape', 'rounded'):
        source = run / f'{shape}.pptx'
        subprocess.run(['node', 'scripts/prepare-notes-size-boundary.mjs', str(source), shape], cwd=ROOT, check=True)
        for round in range(4):
            path = source if round == 0 else run / shape / str(round) / f'{shape}.pptx'
            if round:
                path.parent.mkdir(parents=True)
                previous = source if round == 1 else run / shape / str(round - 1) / f'{shape}.pptx'
                process = subprocess.run([str(SOFFICE), f'-env:UserInstallation={(run / f"profile-{shape}-{round}").as_uri()}', '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', str(path.parent), str(previous)], capture_output=True, text=True, timeout=90)
                logs.append({'scenario': shape, 'round': round, 'exitCode': process.returncode, 'stdout': process.stdout, 'stderr': process.stderr})
                if process.returncode or not path.is_file(): raise RuntimeError(f'LibreOffice conversion failed: {process.stderr}')
            records.append({'scenario': shape, 'round': round, 'file': str(path.relative_to(ROOT)), 'sha256': digest(path), 'inventory': inventory(path)})
    evidence = {'scope': 'Three fresh two-slide decks with schema-boundary dimensions; three LibreOfficeDev headless resaves each, independent OPC/XML inventory only. No GUI, PowerPoint or notes PDF.',
                'application': subprocess.run([str(SOFFICE), '--version'], capture_output=True, text=True, check=True).stdout.strip(), 'platform': platform.platform(),
                'generatorSha256': digest(ROOT / 'scripts/prepare-notes-size-boundary.mjs'), 'runnerSha256': digest(Path(__file__)), 'inventorySha256': digest(ROOT / 'scripts/research/notes_size_inventory.py'), 'records': records, 'logs': logs}
    (ROOT / 'docs/testing/notes-size-libreoffice-evidence.json').write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps({'records': len(records), 'application': evidence['application']}, ensure_ascii=False))

if __name__ == '__main__': main()
