#!/usr/bin/env python3
"""Check retained files and recorded UI state for the corpus text surface observation."""
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
RECORD = ROOT / 'docs/testing/corpus-text-surfaces-powerpoint-evidence.json'


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify(record_path=RECORD, root=ROOT):
    record = json.loads(record_path.read_text())
    directory = root / record['artifactDirectory']
    for name, expected in record['files'].items():
        if digest(directory / name) != expected:
            raise ValueError(f'artifact hash mismatch: {name}')
    saved = root / 'tests/fixtures/corpus-text-powerpoint-interactive.pptx'
    if digest(saved) != record['files']['text-surface.pptx']:
        raise ValueError('surface input differs from retained saved fixture')
    spec = importlib.util.spec_from_file_location('corpus_text', root / 'scripts/research/verify-corpus-text-powerpoint.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.inspect(saved, True)
    notes = (directory / 'text-notes.ax.txt').read_text()
    before = (directory / 'link-before.ax.txt').read_text()
    destination = (directory / 'link-destination.ax.txt').read_text()
    if not all(s in notes for s in ('幻灯片 2 / 2', '隐藏', 'Speaker & reviewer notes')):
        raise ValueError('notes capture does not show hidden slide and expected notes')
    if not all(s in before for s in ('幻灯片 1', 'Go to notes', '幻灯片 2')):
        raise ValueError('starting capture does not show slide 1 and its target')
    if '布局区域 幻灯片 2' not in destination:
        raise ValueError('destination capture does not show slide 2')


if __name__ == '__main__':
    verify()
    print('verified corpus text surface evidence')
