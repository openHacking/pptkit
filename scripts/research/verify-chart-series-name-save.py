#!/usr/bin/env python3
"""Independent oracle for the first series name of the two-series fixture."""
import importlib.util
import json
import sys
from pathlib import Path

spec = importlib.util.spec_from_file_location('label_save_oracle', Path(__file__).with_name('verify-chart-label-save.py'))
oracle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(oracle)

def verify(before, after, expected):
    old, new = oracle.entries(before), oracle.entries(after)
    assert old.keys() == new.keys()
    for name in old.keys() - {oracle.CHART, oracle.BOOK}:
        assert old[name] == new[name], name
    c, s = oracle.C, oracle.S
    oracle.compare_xml(old[oracle.CHART], new[oracle.CHART], [f'.//{c}barChart/{c}ser[1]/{c}tx/{c}strRef/{c}strCache/{c}pt[@idx="0"]/{c}v'], expected)
    old_book, new_book = oracle.entries(old[oracle.BOOK]), oracle.entries(new[oracle.BOOK])
    assert old_book.keys() == new_book.keys()
    for name in old_book.keys() - {oracle.SHEET}:
        assert old_book[name] == new_book[name], name
    oracle.compare_xml(old_book[oracle.SHEET], new_book[oracle.SHEET], [f'.//{s}c[@r="B1"]/{s}is/{s}t'], expected, allow_space=True)

if __name__ == '__main__':
    assert len(sys.argv) == 4, 'usage: verify-chart-series-name-save.py before.pptx after.pptx expected-json'
    verify(Path(sys.argv[1]).read_bytes(), Path(sys.argv[2]).read_bytes(), json.loads(sys.argv[3]))
    print('independent chart series name verification passed')
