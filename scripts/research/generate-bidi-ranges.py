"""Generate the preview-routing table from macOS ICU's Unicode 17 data."""
import ctypes as C
import ctypes.util
from pathlib import Path

lib = C.CDLL(ctypes.util.find_library('icucore'))
lib.u_getUnicodeVersion.argtypes = [C.POINTER(C.c_uint8)]
version = (C.c_uint8 * 4)()
lib.u_getUnicodeVersion(version)
if list(version) != [17, 0, 0, 0]:
    raise SystemExit(f'Expected Unicode 17.0.0.0, found {list(version)}')
lib.u_charDirection.argtypes = [C.c_int32]
lib.u_charDirection.restype = C.c_int8
ranges = []
for codepoint in range(0x110000):
    # ICU UCharDirection: R = 1, AL = 13, AN = 5. Include assigned and
    # default bidi classes in unassigned ranges, as returned by ICU.
    if lib.u_charDirection(codepoint) in (1, 5, 13):
        if ranges and ranges[-1][1] + 1 == codepoint:
            ranges[-1][1] = codepoint
        else:
            ranges.append([codepoint, codepoint])
lines = ['  ' + ', '.join(f'0x{value:x}' for pair in ranges[i:i + 4] for value in pair) + ',' for i in range(0, len(ranges), 4)]
source = '''// Generated from macOS ICU Unicode 17.0 bidi classes R, AL and AN, including
// default classes and ICU private-use values. See scripts/research/generate-bidi-ranges.py.
// This is a routing predicate, not a Unicode bidi implementation.
const ranges: readonly number[] = [\n''' + '\n'.join(lines) + '''\n];

/** Whether paragraph layout needs direction-sensitive browser handling. */
export function needsBidiParagraphLayout(text: string): boolean {
  for (const character of text) {
    const code = character.codePointAt(0)!;
    if (code === 0x200e || (code >= 0x202a && code <= 0x202e) || (code >= 0x2066 && code <= 0x2069)) return true;
    let low = 0, high = ranges.length / 2;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (code < ranges[middle * 2]!) high = middle;
      else if (code > ranges[middle * 2 + 1]!) low = middle + 1;
      else return true;
    }
  }
  return false;
}
'''
path = Path('packages/svg-renderer/src/rendering/bidi-ranges.ts')
path.write_text(source)
print(f'{len(ranges)} ranges written to {path}')
