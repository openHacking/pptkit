#!/usr/bin/env python3
"""Check one notes text edit against actual ZIP payloads, independent of parser internals."""
import sys
import zipfile

before, after, part, old, new = sys.argv[1:]
with zipfile.ZipFile(before) as source, zipfile.ZipFile(after) as candidate:
    left, right = source.namelist(), candidate.namelist()
    if len(left) != len(set(left)) or len(right) != len(set(right)) or set(left) != set(right): raise ValueError('part inventory changed')
    changed = [name for name in left if source.read(name) != candidate.read(name)]
    if changed != [part]: raise ValueError('unexpected changed parts')
    original = source.read(part); replaced = original.replace(old.encode(), new.encode())
    if original.count(old.encode()) != 1 or replaced != candidate.read(part): raise ValueError('notes edit changed more than the selected text')
print('verified')
