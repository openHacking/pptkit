"""Check compiled routing data against Python's independent Unicode database."""
import base64
import ctypes as C
import ctypes.util
import json
import subprocess
import unicodedata as U

assert U.unidata_version == '16.0.0', U.unidata_version
script = r'''
import { needsBidiParagraphLayout } from './packages/svg-renderer/dist/rendering/bidi-ranges.js';
const bits = Buffer.alloc(Math.ceil(0x110000 / 8));
for (let code = 0; code < 0x110000; code++) if (needsBidiParagraphLayout(String.fromCodePoint(code))) bits[code >>> 3] |= 1 << (code & 7);
process.stdout.write(bits.toString('base64'));
'''
bits = base64.b64decode(subprocess.check_output(['node', '--input-type=module', '-e', script]))
checked = rtl = skipped = 0
mismatches = []
for code in range(0x110000):
    char = chr(code)
    if U.category(char) == 'Cn':
        skipped += 1
        continue
    expected = U.bidirectional(char) in ('R', 'AL', 'AN') or code == 0x200e or 0x202a <= code <= 0x202e or 0x2066 <= code <= 0x2069
    actual = bool(bits[code >> 3] & (1 << (code & 7)))
    checked += 1
    rtl += expected
    if expected != actual:
        mismatches.append({'codepoint': f'U+{code:04X}', 'category': U.category(char), 'bidi': U.bidirectional(char), 'expected': expected, 'actual': actual})
print(json.dumps({'pythonUnicode': U.unidata_version, 'checkedAssigned': checked, 'matchedDirectionSensitive': rtl, 'skippedUnassigned': skipped, 'conservativeDifferences': mismatches}, indent=2))
# macOS ICU 17 classifies these private-use values as direction-sensitive;
# Python Unicode 16 classifies them as L. Retaining the conservative routing
# avoids assuming a user's private glyphs are left-to-right.
known_private = {'U+F882', 'U+F883', 'U+F89A', 'U+F89B', 'U+F89C', 'U+F89D', 'U+F89E'}
assert {item['codepoint'] for item in mismatches} == known_private
assert all(item['category'] == 'Co' and item['actual'] and not item['expected'] for item in mismatches)

lib = C.CDLL(ctypes.util.find_library('icucore'))
lib.u_getUnicodeVersion.argtypes = [C.POINTER(C.c_uint8)]
version = (C.c_uint8 * 4)()
lib.u_getUnicodeVersion(version)
assert list(version) == [17, 0, 0, 0], list(version)
lib.u_charDirection.argtypes = [C.c_int32]
lib.u_charDirection.restype = C.c_int8
for code in range(0x110000):
    expected = lib.u_charDirection(code) in (1, 5, 13) or code == 0x200e or 0x202a <= code <= 0x202e or 0x2066 <= code <= 0x2069
    actual = bool(bits[code >> 3] & (1 << (code & 7)))
    assert expected == actual, f'ICU 17 mismatch at U+{code:04X}'
print('macOS ICU Unicode 17: all 1,114,112 code points agree')
