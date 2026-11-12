"""Check a retained physical-series swap without trusting PPTKit's parser."""
import re
import sys
import zipfile

before, after = sys.argv[1:]
chart = 'ppt/charts/chart1.xml'
with zipfile.ZipFile(before) as source, zipfile.ZipFile(after) as target:
    assert set(source.namelist()) == set(target.namelist())
    for name in source.namelist():
        if name != chart:
            assert source.read(name) == target.read(name), name
    old = source.read(chart)
    found = list(re.finditer(rb'<c:ser>.*?</c:ser>', old, re.S))
    assert len(found) == 2
    a, b = found
    expected = old[:a.start()] + b.group() + old[a.end():b.start()] + a.group() + old[b.end():]
    assert expected == target.read(chart)
