#!/usr/bin/env python3
"""Check a previously captured PowerPoint group move; never drives UI."""
import importlib.util
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location('edit_native', Path(__file__).with_name('verify-edit-operation-native.py'))
native = importlib.util.module_from_spec(spec)
spec.loader.exec_module(native)

def verify(local=False):
    record = json.loads((ROOT / 'docs/testing/edit-operation-move-powerpoint-evidence.json').read_text())
    for item in record['artifacts'].values():
        if native.sha(ROOT / item['path']) != item['sha256']: raise ValueError('artifact hash differs')
    before = native.inspect(ROOT / record['artifacts']['input']['path'], 'Edited native')
    after = native.inspect(ROOT / record['artifacts']['output']['path'], 'Edited native', ('2345634', '1974573'))
    if before != after: raise ValueError('headline or child geometry differs')
    if local:
        directory = ROOT / record['artifactDirectory']
        for name, value in record['captures'].items():
            if native.sha(directory / name) != value: raise ValueError('capture hash differs')
        for name in ('move-selected.ax.txt', 'move-after.ax.txt'):
            state = (directory / name).read_text()
            if 'Selected:\n\t21 组合的对象 (disabled) group' not in state: raise ValueError('group selection missing')
        state = (directory / 'move-reopen.ax.txt').read_text()
        if any(value not in state for value in ('native-move', 'Edited native', 'Edited shape', '组合的对象 (disabled) group')): raise ValueError('reopened state differs')

if __name__ == '__main__':
    if sys.argv[1:] not in ([], ['--local']): raise SystemExit('usage: verify-edit-operation-move.py [--local]')
    verify(bool(sys.argv[1:]))
    print('verified group move' + (' and local captures' if sys.argv[1:] else ''))
