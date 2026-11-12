/** Bounded, standalone Huffman DCT JPEG structure; entropy data is not decoded. */
export function validateJpegStructure(input: Uint8Array, signal?: AbortSignal): { precision: number } {
  const fail = (reason: string): never => { throw new Error(`Invalid or unsupported JPEG: ${reason}`); };
  const tag = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(Uint8Array.prototype), Symbol.toStringTag)?.get;
  if (!ArrayBuffer.isView(input) || tag?.call(input) !== "Uint8Array") throw new TypeError("JPEG structure input must be a Uint8Array.");
  if (input.byteLength > 64 * 1024 * 1024) throw new RangeError("JPEG structure exceeds the 64 MiB limit.");
  signal?.throwIfAborted();
  const bytes = new Uint8Array(input);
  if (bytes[0] !== 0xff || bytes[1] !== 0xd8) fail("missing SOI");
  const word = (i: number) => (bytes[i]! << 8) | bytes[i + 1]!;
  const quant = new Map<number, number>(), huffman = new Set<number>();
  const components = new Map<number, { quant: number; bands: Int8Array }>();
  let at = 2, markers = 0, process = 0, precision = 0, scans = 0, interval = 0, restart = 0, scanning = false, data = false;
  while (at < bytes.length) {
    signal?.throwIfAborted();
    if (scanning && bytes[at] !== 0xff) {
      const stop = Math.min(bytes.length, at + 32768);
      while (at < stop && bytes[at] !== 0xff) at++;
      data = true;
      continue;
    }
    if (bytes[at++] !== 0xff) fail("expected marker");
    const start = at;
    while (bytes[at] === 0xff) { if ((at & 32767) === 0) signal?.throwIfAborted(); at++; }
    if (at === bytes.length) fail("truncated marker");
    const marker = bytes[at++]!;
    if (scanning && marker === 0) {
      if (at !== start + 1) fail("invalid scan byte stuffing");
      data = true; continue;
    }
    if (++markers > 100000) throw new RangeError("JPEG structure exceeds the 100,000 marker limit.");
    if (scanning && marker >= 0xd0 && marker <= 0xd7) {
      if (!interval || !data || marker !== 0xd0 + restart) fail("invalid restart marker sequence");
      restart = (restart + 1) & 7; data = false;
      continue;
    }
    if (scanning && !data) fail("empty scan segment");
    scanning = false;
    if (marker === 0xd9) {
      if (!scans || at !== bytes.length || [...components.values()].some(component => component.bands[0] === -1)) fail("incomplete image or bytes after EOI");
      return { precision };
    }
    if (at + 2 > bytes.length) fail("missing segment length");
    const length = word(at);
    if (length < 2 || length > bytes.length - at) fail("truncated marker segment");
    const end = at + length;
    let p = at + 2;
    if (marker >= 0xe0 && marker <= 0xef || marker === 0xfe) {
      // Application metadata and comments are opaque to this check.
    } else if (marker === 0xdb) {
      if (p === end) fail("empty quantization segment");
      while (p < end) {
        const selector = bytes[p++]!, size = (selector >>> 4) + 1, id = selector & 15;
        if (size > 2 || id > 3 || end - p < 64 * size) fail("invalid quantization table");
        for (let i = 0; i < 64; i++, p += size) if (!(size === 1 ? bytes[p] : word(p))) fail("zero quantization value");
        quant.set(id, size);
      }
    } else if (marker === 0xc4) {
      if (p === end) fail("empty Huffman segment");
      while (p < end) {
        const selector = bytes[p++]!;
        if ((selector >>> 4) > 1 || (selector & 15) > 3 || end - p < 16) fail("invalid Huffman table header");
        let slots = 1, symbols = 0;
        for (let i = 0; i < 16; i++) {
          const count = bytes[p++]!;
          slots = slots * 2 - count; symbols += count;
          if (slots <= 0) fail("Huffman table has no reserved all-ones code");
        }
        if (!symbols || symbols > 256 || symbols > end - p) fail("invalid Huffman symbol count");
        p += symbols; huffman.add(selector);
      }
    } else if (marker === 0xdd) {
      if (length !== 4) fail("invalid restart interval segment");
      interval = word(p);
    } else if (marker === 0xc0 || marker === 0xc1 || marker === 0xc2) {
      if (process || length < 11 || !word(p + 1) || !word(p + 3)) fail("invalid or duplicate frame");
      precision = bytes[p]!; const count = bytes[p + 5]!;
      if ((precision !== 8 && (marker === 0xc0 || precision !== 12)) || !count || length !== 8 + 3 * count) fail("unsupported frame fields");
      for (let i = 0; i < count; i++) {
        const id = bytes[p + 6 + i * 3]!, sampling = bytes[p + 7 + i * 3]!, table = bytes[p + 8 + i * 3]!;
        const h = sampling >>> 4, v = sampling & 15;
        if (components.has(id) || h < 1 || h > 4 || v < 1 || v > 4 || table > 3) fail("invalid frame component");
        components.set(id, { quant: table, bands: new Int8Array(64).fill(-1) });
      }
      process = marker;
    } else if (marker === 0xda) {
      if (!process || length < 8) fail("scan before frame or truncated scan header");
      const count = bytes[p]!;
      if (!count || count > 4 || length !== 6 + 2 * count) fail("invalid scan component count");
      const ss = bytes[end - 3]!, se = bytes[end - 2]!, ah = bytes[end - 1]! >>> 4, al = bytes[end - 1]! & 15;
      const progressive = process === 0xc2;
      if (progressive ? ss > se || se > 63 || (ss === 0 ? se !== 0 : count !== 1) || ah > 13 || al > 13 || (ah !== 0 && ah !== al + 1)
        : ss !== 0 || se !== 63 || ah !== 0 || al !== 0) fail("invalid scan spectral or approximation fields");
      const ids = new Set<number>(), limit = process === 0xc0 ? 1 : 3;
      for (let i = 0; i < count; i++) {
        const id = bytes[p + 1 + i * 2]!, selector = bytes[p + 2 + i * 2]!;
        const component = components.get(id) ?? fail("unknown scan component");
        if (ids.has(id)) fail("duplicate scan component");
        ids.add(id);
        const quantSize = quant.get(component.quant);
        if (!quantSize || (process === 0xc0 && quantSize !== 1)) fail("missing or unsupported quantization table");
        const dc = selector >>> 4, ac = selector & 15;
        if (dc > limit || ac > limit || ((!progressive || (ss === 0 && ah === 0)) && !huffman.has(dc)) ||
          ((!progressive || ss !== 0) && !huffman.has(0x10 | ac))) fail("missing or unsupported Huffman table");
        if (ss > 0 && component.bands[0] === -1) fail("AC scan precedes DC scan");
        for (let band = ss; band <= se; band++) {
          const previous = component.bands[band]!;
          if (previous === -1 ? ah !== 0 : ah === 0 || ah !== previous) fail("duplicate or out-of-order coefficient scan");
          component.bands[band] = al;
        }
      }
      scans++; scanning = true; data = false; restart = 0;
    } else fail(`marker FF${marker.toString(16).padStart(2, "0")} is outside the supported profile`);
    at = end;
  }
  return fail("missing EOI");
}
