"""Independently verify a two-series order-only save, including byte scope."""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

before, after = sys.argv[1:]
chart_name = 'ppt/charts/chart1.xml'
ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
with zipfile.ZipFile(before) as original, zipfile.ZipFile(after) as changed:
    assert set(original.namelist()) == set(changed.namelist())
    for name in original.namelist():
        if name != chart_name:
            assert original.read(name) == changed.read(name), name
    old, new = original.read(chart_name), changed.read(chart_name)
    def mapping(data):
        root = ET.fromstring(data)
        return {int(ser.find('c:idx', ns).attrib['val']): int(ser.find('c:order', ns).attrib['val']) for ser in root.findall('.//c:ser', ns)}
    old_mapping, new_mapping = mapping(old), mapping(new)
    assert len(old_mapping) == len(new_mapping) == 2
    assert set(old_mapping) == set(new_mapping)
    ids = list(old_mapping)
    assert new_mapping[ids[0]] == old_mapping[ids[1]] and new_mapping[ids[1]] == old_mapping[ids[0]]
    # The fixture's direct order attributes are the only byte changes in chart XML.
    pattern = rb'(<c:order val=")[0-9]+("/>)'
    assert len(re.findall(pattern, old)) == len(re.findall(pattern, new)) == 2
    assert re.sub(pattern, rb'\1#\2', old) == re.sub(pattern, rb'\1#\2', new)
