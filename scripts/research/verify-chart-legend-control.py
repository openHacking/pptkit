"""Check a control differs from its base only by one indexed legend entry."""
import sys
import zipfile

base_path, control_path = sys.argv[1:]
chart = 'ppt/charts/chart1.xml'
needle = b'<c:legend><c:legendPos val="r"/>'
addition = b'<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>'
with zipfile.ZipFile(base_path) as base, zipfile.ZipFile(control_path) as control:
    assert set(base.namelist()) == set(control.namelist())
    for name in base.namelist():
        if name == chart:
            original = base.read(name)
            assert original.count(needle) == 1
            assert original.replace(needle, needle + addition, 1) == control.read(name)
        else:
            assert base.read(name) == control.read(name), name
