#!/usr/bin/env python3
"""Re-read retained boundary PPTX files and verify the recorded scoped observations."""
import hashlib
import json
import sys
from pathlib import Path
from notes_size_inventory import inventory

ROOT = Path(__file__).resolve().parents[2]
BASE = (ROOT / '.tmp/notes-size-libreoffice').resolve()
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()

def verify(path):
    evidence = json.loads(path.read_text())
    for key, file in [('generatorSha256', 'scripts/prepare-notes-size-boundary.mjs'), ('runnerSha256', 'scripts/research/notes-size-libreoffice.py'), ('inventorySha256', 'scripts/research/notes_size_inventory.py')]:
        if evidence[key] != sha(ROOT / file): raise ValueError('source dependency changed')
    records = evidence['records']; logs = evidence['logs']
    if len(records) != 12 or len(logs) != 9: raise ValueError('invalid observation count')
    outcomes = {}
    for index, shape in enumerate(['portrait', 'landscape', 'rounded']):
        frames = []
        for round in range(4):
            record = records[index * 4 + round]
            if record['scenario'] != shape or record['round'] != round: raise ValueError('observation order changed')
            file = (ROOT / record['file']).resolve()
            if not file.is_relative_to(BASE) or sha(file) != record['sha256']: raise ValueError('file identity changed')
            actual = inventory(file)
            if actual != record['inventory']: raise ValueError('inventory disagrees with actual file')
            expected_size = [51206400, 914400] if shape == 'landscape' else [914400, 51206400]
            if actual['slideSize'] != expected_size or actual['notesSize'] != [6858000, 9144000] or len(actual['slides']) != 2: raise ValueError('unexpected slide dimensions or page count')
            for i, suffix in enumerate(['FIRST', 'SECOND']):
                slide = actual['slides'][i]
                if slide['slideText'] != f'SLIDE-{shape}-{suffix}' or slide['notesText'] != f'NOTES-{shape}-{suffix}': raise ValueError('page text or order changed')
                if round == 0 and slide['localImage']['frame'] is not None: raise ValueError('source local geometry unexpectedly overrides master')
                if round and (slide['localImage']['preset'] != 'rect' or slide['localImage']['frame'] is None): raise ValueError('native local image geometry missing')
            if actual['masterImage']['preset'] != 'rect': raise ValueError('master image geometry missing')
            frame = actual['masterImage']['frame'] if round == 0 else actual['slides'][0]['localImage']['frame']
            if round == 0:
                sw, sh = actual['slideSize']
                if abs(frame['width'] * sh - frame['height'] * sw) > max(sw, sh) or abs(2 * frame['x'] + frame['width'] - 6858000) > 1 or abs(2 * frame['y'] + frame['height'] - 4572000) > 1: raise ValueError('source notes image fit is invalid')
            if round and any(slide['localImage']['frame'] != frame for slide in actual['slides']): raise ValueError('local image frames differ')
            frames.append(frame)
        for prior, current in zip(frames[1:], frames[2:]):
            if current['x'] != prior['x'] or current['y'] != prior['y'] or current['width'] != prior['width'] - 360 or current['height'] != prior['height'] - 360: raise ValueError('observed native drift differs')
        outcomes[shape] = frames
    if [(x['scenario'], x['round']) for x in logs] != [(shape, round) for shape in ['portrait', 'landscape', 'rounded'] for round in range(1, 4)] or any(x['exitCode'] != 0 for x in logs): raise ValueError('native conversion failure or log identity changed')
    return {'verified': True, 'records': len(records), 'slides': 24, 'nativeRounds': len(logs), 'frames': outcomes}

if __name__ == '__main__': print(json.dumps(verify(Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'docs/testing/notes-size-libreoffice-evidence.json'), ensure_ascii=False))
