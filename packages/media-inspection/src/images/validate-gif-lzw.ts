/** Verify index counts and palette references without expanding the index raster. */
export function createGifLzwValidator(minimum: number, paletteEntries: number, expected: number) {
  const clear = 1 << minimum, end = clear + 1;
  const lengths = new Uint16Array(4096);
  let next = end + 1, width = minimum + 1, previous = -1;
  let bits = 0, bitCount = 0, pixels = 0, finished = false;
  const fail = (message: string): never => { throw new Error(`Invalid GIF LZW: ${message}`); };
  const lengthOf = (code: number) => code < clear ? 1 : lengths[code]!;

  return {
    push(bytes: Uint8Array, start: number, size: number) {
      for (let i = start; i < start + size; i++) {
        if (finished) fail("image bytes after end code");
        bits |= bytes[i]! << bitCount;
        bitCount += 8;
        while (bitCount >= width) {
          const code = bits & ((1 << width) - 1);
          bits >>>= width;
          bitCount -= width;
          if (code === clear) {
            next = end + 1; width = minimum + 1; previous = -1;
            continue;
          }
          if (code === end) {
            if (pixels !== expected) fail("index count does not match frame dimensions");
            finished = true;
            break; // The unused bits of this last byte are padding.
          }
          if (code < clear && code >= paletteEntries) fail("palette index is out of range");
          if (code > next || (code === next && (previous < 0 || next >= 4096))) fail("invalid dictionary code");
          const repeated = code === next;
          pixels += repeated ? lengthOf(previous) + 1 : lengthOf(code);
          if (pixels > expected) fail("index count exceeds frame dimensions");
          if (previous >= 0 && next < 4096) {
            lengths[next] = lengthOf(previous) + 1;
            next++;
            if (next === (1 << width) && width < 12) width++;
          }
          previous = code;
        }
      }
    },
    finish() {
      if (!finished) fail("missing end code");
    },
  };
}
