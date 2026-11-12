"""Compare the recorded Chromium DOM positions with the system ICU bidi algorithm."""
import ctypes
import ctypes.util
import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

folder = Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/bidi-observation-v3')
expected = {
    'hebrew-single': ('ABC (שלום 123) XYZ', 1),
    'hebrew-runs': ('ABC (שלום 123) XYZ', 4),
    'arabic-single': ('ABC (مرحبا 123) XYZ', 1),
    'arabic-runs': ('ABC (مرحبا 123) XYZ', 4),
}
ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
with zipfile.ZipFile(folder / 'source.pptx') as archive:
    root = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
    shapes = {}
    for shape in root.findall('.//p:sp', ns):
        name = shape.find('./p:nvSpPr/p:cNvPr', ns)
        if name is not None:
            runs = shape.findall('.//a:r', ns)
            shapes[name.get('name')] = (''.join(run.findtext('a:t', default='', namespaces=ns) for run in runs), len(runs))
    assert shapes == expected, shapes

lib = ctypes.CDLL(ctypes.util.find_library('icucore'))
lib.ubidi_open.restype = ctypes.c_void_p
lib.ubidi_setPara.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_uint16), ctypes.c_int32,
                             ctypes.c_uint8, ctypes.POINTER(ctypes.c_uint8), ctypes.POINTER(ctypes.c_int32)]
lib.ubidi_getVisualIndex.argtypes = [ctypes.c_void_p, ctypes.c_int32, ctypes.POINTER(ctypes.c_int32)]
lib.ubidi_getVisualIndex.restype = ctypes.c_int32
lib.ubidi_close.argtypes = [ctypes.c_void_p]
rows = json.loads((folder / 'positions.json').read_text())
assert len(rows) == len(expected) and {row['id'] for row in rows} == set(expected)
for row in rows:
    chars = row['chars']
    text = ''.join(char['value'] for char in chars)
    assert text == expected[row['id']][0]
    assert all(len(c) == 1 and ord(c) <= 0xffff for c in text)
    assert [char['index'] for char in chars] == list(range(len(text)))
    assert all(char['width'] > 0 for char in chars)
    buffer = (ctypes.c_uint16 * len(text))(*map(ord, text))
    error = ctypes.c_int32(0)
    bidi = lib.ubidi_open()
    assert bidi
    try:
        lib.ubidi_setPara(bidi, buffer, len(text), 0, None, ctypes.byref(error))
        indices = [lib.ubidi_getVisualIndex(bidi, i, ctypes.byref(error)) for i in range(len(text))]
        assert error.value == 0, error.value
    finally:
        lib.ubidi_close(bidi)
    assert sorted(indices) == list(range(len(text))), indices
    visual = sorted(range(len(text)), key=indices.__getitem__)
    actual = [char['index'] for char in sorted(chars, key=lambda char: char['x'])]
    assert actual == visual, (row['id'], actual, visual)
    print(f"{row['id']}: {len(chars)} character positions match ICU")
