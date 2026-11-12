/** Bounded JPEG frame-precision probe for emitted data URLs; not a JPEG validator. */
export function jpegDataUrlPrecision(url: string): number | "not-jpeg" | "unavailable" {
  const comma = url.indexOf(",");
  if (comma < 0 || !/^data:image\//i.test(url)) return "unavailable";
  const metadata = url.slice(5, comma), declaredJpeg = /^image\/(?:jpe?g|pjpeg)(?:;|$)/i.test(metadata);
  const max = 65536, source = url.slice(comma + 1);
  let bytes: Uint8Array;
  if (/;base64$/i.test(metadata)) {
    const prefix = source.slice(0, Math.ceil(max / 3) * 4);
    // Deliberately fail closed on escaped or whitespace-bearing Base64.
    if (!/^[A-Za-z0-9+/]*={0,2}$/.test(prefix)) return "unavailable";
    try { const decoded = atob(prefix); bytes = Uint8Array.from(decoded.slice(0, max), char => char.charCodeAt(0)); }
    catch { return "unavailable"; }
  } else {
    const values: number[] = [];
    for (let i = 0; i < source.length && values.length < max; i++) {
      const code = source.charCodeAt(i);
      if (code === 37) {
        const digits = source.slice(i + 1, i + 3);
        if (!/^[0-9a-f]{2}$/i.test(digits)) return "unavailable";
        values.push(Number.parseInt(digits, 16)); i += 2;
      } else {
        if (code > 127 || code === 35) return "unavailable";
        values.push(code);
      }
    }
    bytes = Uint8Array.from(values);
  }
  if (bytes.length < 2 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return declaredJpeg ? "unavailable" : "not-jpeg";
  let position = 2;
  while (position < bytes.length) {
    if (bytes[position++] !== 0xff) return "unavailable";
    while (position < bytes.length && bytes[position] === 0xff) position++;
    if (position >= bytes.length) return "unavailable";
    const marker = bytes[position++]!;
    if (marker === 0 || marker === 0xd8 || marker === 0xd9 || marker === 0xda) return "unavailable";
    if (marker === 1 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (position + 2 > bytes.length) return "unavailable";
    const length = bytes[position]! * 256 + bytes[position + 1]!;
    if (length < 2 || length > bytes.length - position) return "unavailable";
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) return length >= 8 ? bytes[position + 2]! : "unavailable";
    position += length;
  }
  return "unavailable";
}
