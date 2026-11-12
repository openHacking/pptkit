"""Revalidate retained actual ZIP/PDF counterexamples without overwriting them."""
import hashlib
import importlib.util
import json
from datetime import datetime, timezone
from pathlib import Path


def load(name):
    spec = importlib.util.spec_from_file_location(name, Path(__file__).with_name('native-notes-' + name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def sha(file):
    return hashlib.sha256(file.read_bytes()).hexdigest()


def verify():
    append, styles, themes = (load(name) for name in ['append', 'styles', 'themes'])
    records = []
    for stage in ['append', 'styles', 'themes']:
        manifest_path = Path('docs/testing/native-notes-' + stage + '-evidence.json')
        manifest = json.loads(manifest_path.read_text())
        for group in manifest['groups']:
            name = group.get('variant', group.get('group'))
            directory = Path(manifest['root']) / name
            expected = json.loads((directory / 'expected.json').read_text())
            inspection = directory / {'append': 'append-negative-inspection.json', 'styles': 'styles-negative-inspection.json', 'themes': 'theme-negative-inspection.json'}[stage]
            negatives = json.loads(inspection.read_text())
            if stage == 'append':
                append.inspect_package(directory / 'append-round1.pptx', directory, expected)
                append.inspect_pdf(directory / 'append-round1-notes.pdf', directory, expected)
            elif stage == 'styles':
                for operation in styles.STAGES:
                    styles.inspect_package(directory / (operation + '-round1.pptx'), directory, expected, operation)
                    styles.inspect_pdf(directory / (operation + '-round1-notes.pdf'), directory, expected, operation)
            else:
                controls = json.loads((directory / 'controls.json').read_text()) if name == 'override' else [{**c, 'variant': name} for c in expected['controls']]
                themes.inspect_theme_package(directory / 'office-round1.pptx', directory, expected, controls)
                baseline = directory / ('native-baseline-round3-notes.pdf' if name == 'override' else 'baseline-notes.pdf')
                themes.inspect_pdf(directory / 'office-round1-notes.pdf', baseline, controls, True)
            for control in negatives['controls']:
                file = Path(control['file'])
                assert sha(file) == control['sha256'], ('Captured negative bytes changed', file)
                try:
                    if stage == 'append':
                        if file.suffix == '.pptx':
                            append.inspect_package(file, directory, expected)
                        else:
                            reference = directory / 'append-round1-notes.pdf' if control['name'] == 'later-round' else None
                            append.inspect_pdf(file, directory, expected, reference)
                    elif stage == 'styles':
                        operation = control['stage']
                        if file.suffix == '.pptx':
                            styles.inspect_package(file, directory, expected, operation)
                        else:
                            reference = directory / (operation + '-round1-notes.pdf') if control['name'] == 'later-ink' else None
                            styles.inspect_pdf(file, directory, expected, operation, reference)
                    elif file.suffix == '.pptx':
                        themes.inspect_theme_package(file, directory, expected, controls)
                    else:
                        reference = directory / 'office-round1-notes.pdf' if control['name'] == 'later-ink' else None
                        themes.inspect_pdf(file, baseline, controls, True, reference)
                except (AssertionError, KeyError) as error:
                    records.append({'stage': stage, 'group': name, 'name': control['name'],
                                    'file': str(file), 'sha256': sha(file),
                                    'rejectedBy': type(error).__name__, 'reason': str(error)})
                else:
                    raise AssertionError('Captured corruption accepted: ' + str(file))
    assert len(records) == 143
    result = {'schemaVersion': 1, 'recordedAt': datetime.now(timezone.utc).isoformat(),
              'total': {'capturedNegativeFilesRejected': len(records)}, 'records': records,
              'scope': 'Actual retained ZIP/PDF bytes hash-checked and rejected by current independent oracles; positive stage controls checked first. No native actions or regenerated counterexamples.'}
    Path('docs/testing/native-notes-captured-negatives-evidence.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps(result['total']))


if __name__ == '__main__':
    verify()
