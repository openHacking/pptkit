"""Check the controlled move of two PowerPoint-produced series, byte for byte."""
import re
import sys
import zipfile

before_path, after_path = sys.argv[1:]
chart = 'ppt/charts/chart1.xml'
series_pattern = rb'<c:ser>.*?</c:ser>'
order_pattern = rb'(<c:order val=")[0-9]+("/>)'
legend_pattern = rb'(<c:legendEntry>\s*<c:idx val=")[0-9]+("/>)'
uid_pattern = rb'<c16:uniqueId\s+val="([^"]+)"'

with zipfile.ZipFile(before_path) as original, zipfile.ZipFile(after_path) as candidate:
    assert set(original.namelist()) == set(candidate.namelist())
    for name in original.namelist():
        if name != chart:
            assert original.read(name) == candidate.read(name), name
    old, new = original.read(chart), candidate.read(chart)
    old_series, new_series = re.findall(series_pattern, old, re.S), re.findall(series_pattern, new, re.S)
    assert len(old_series) == len(new_series) == 2
    def identity(data):
        value = re.search(rb'<c:idx val="([0-9]+)"/>', data)
        assert value
        return value.group(1)
    assert [identity(item) for item in new_series] == list(reversed([identity(item) for item in old_series]))
    def normalize(data):
        assert len(re.findall(order_pattern, data)) == 1
        return re.sub(order_pattern, rb'\1#\2', data)
    assert [normalize(item) for item in new_series] == [normalize(item) for item in reversed(old_series)]
    assert re.search(legend_pattern, old).group(0).endswith(b'1"/>')
    assert re.search(legend_pattern, new).group(0).endswith(b'0"/>')
    def surroundings(data):
        data = re.sub(series_pattern, b'<SERIES/>', data, flags=re.S)
        return re.sub(legend_pattern, rb'\1#\2', data)
    assert surroundings(old) == surroundings(new)
    for data in old_series:
        ids = re.findall(uid_pattern, data)
        assert len(ids) == 1
        for uid in ids:
            assert sum(original.read(name).count(uid) for name in original.namelist()) == 1
