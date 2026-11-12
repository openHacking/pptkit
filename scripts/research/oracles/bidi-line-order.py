"""Compare each captured browser line to ICU's paragraph-resolved line order."""
import ctypes as C
import ctypes.util
import json
import sys
import unicodedata
from pathlib import Path

lib = C.CDLL(ctypes.util.find_library('icucore'))
lib.ubidi_open.restype = C.c_void_p
lib.ubidi_close.argtypes = [C.c_void_p]
lib.ubidi_setPara.argtypes = [C.c_void_p, C.POINTER(C.c_uint16), C.c_int32, C.c_uint8, C.POINTER(C.c_uint8), C.POINTER(C.c_int32)]
lib.ubidi_setLine.argtypes = [C.c_void_p, C.c_int32, C.c_int32, C.c_void_p, C.POINTER(C.c_int32)]
lib.ubidi_getVisualIndex.argtypes = [C.c_void_p, C.c_int32, C.POINTER(C.c_int32)]
lib.ubidi_getVisualIndex.restype = C.c_int32
payload = json.loads(Path(sys.argv[1] if len(sys.argv) > 1 else '.tmp/bidi-line-order.json').read_text())
differences = []
ambiguous = []
for row in payload['rows']:
    text = row['text']
    assert all(ord(char) <= 0xffff for char in text)
    chars = [char for line in row['lines'] for char in line]
    assert ''.join(char['char'] for char in chars) == text
    assert [char['index'] for char in chars] == list(range(len(text)))
    array = (C.c_uint16 * len(text))(*map(ord, text))
    error = C.c_int32(0)
    para, line_bidi = lib.ubidi_open(), lib.ubidi_open()
    assert para and line_bidi
    try:
        lib.ubidi_setPara(para, array, len(text), 0, None, C.byref(error))
        for line in row['lines']:
            if not line:
                continue
            start, limit = line[0]['index'], line[-1]['index'] + 1
            lib.ubidi_setLine(para, start, limit, line_bidi, C.byref(error))
            visible = [char for char in line if not char['char'].isspace() and char['width'] > 0]
            # WebKit can return the union of distant bidi fragments for a
            # single-character Range. Such a rectangle is not a position oracle.
            # These fixtures use 24px Arial and no glyph spans 36px. Combining
            # marks can overlap their base and also cannot be sorted this way.
            if any(char['width'] > 36 or unicodedata.category(char['char']).startswith('M') for char in visible):
                ambiguous.append({'id': row['id'], 'logical': ''.join(c['char'] for c in line)})
                continue
            expected = sorted(visible, key=lambda char: lib.ubidi_getVisualIndex(line_bidi, char['index'] - start, C.byref(error)))
            actual = sorted(visible, key=lambda char: char['x'])
            assert error.value == 0, error.value
            if [c['index'] for c in actual] != [c['index'] for c in expected]:
                differences.append({'id': row['id'], 'logical': ''.join(c['char'] for c in line), 'expected': ''.join(c['char'] for c in expected), 'actual': ''.join(c['char'] for c in actual), 'expectedIndices': [c['index'] for c in expected], 'actualIndices': [c['index'] for c in actual]})
    finally:
        lib.ubidi_close(line_bidi)
        lib.ubidi_close(para)
print(json.dumps({'browser': payload['browser'], 'cases': len(payload['rows']), 'ambiguous': ambiguous, 'differences': differences}, ensure_ascii=False, indent=2))
