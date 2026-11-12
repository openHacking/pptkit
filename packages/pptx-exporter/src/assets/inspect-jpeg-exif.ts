import { captureImageInput } from "./capture-image-input.js";

/** Bounded pre-first-scan EXIF orientation inventory; no image decoding or transforms. */
export interface JpegExifOrientationInspection {
  /** No EXIF, EXIF without orientation, or an explicit 0th-IFD orientation. */
  status: "no-exif" | "no-orientation" | "orientation";
  orientation?: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8;
  /** The parser stops at the first scan header or EOI; later metadata is not inspected. */
  stoppedAt: "scan" | "end";
}
export interface JpegExifInspectionOptions {
  maxBytes?: number;
  maxSegments?: number;
  maxIfdEntries?: number;
  signal?: AbortSignal;
}

export function inspectJpegExifOrientation(input: Uint8Array, options: JpegExifInspectionOptions = {}): JpegExifOrientationInspection {
  const limit = (value: number | undefined, fallback: number, name: string): number => {
    const result = value === undefined ? fallback : value;
    if (!Number.isSafeInteger(result) || result <= 0) throw new RangeError(`${name} must be a positive safe integer.`);
    return result;
  };
  const maxBytes = limit(options.maxBytes, 64 * 1024 * 1024, "maxBytes");
  const maxSegments = limit(options.maxSegments, 100000, "maxSegments");
  const maxIfdEntries = limit(options.maxIfdEntries, 65535, "maxIfdEntries");
  const bytes = captureImageInput(input, "JPEG", maxBytes, options.signal);
  const abort = () => { if (options.signal?.aborted) throw options.signal.reason ?? new Error("JPEG inspection aborted."); };
  const fail = (message: string): never => { throw new Error(`Invalid JPEG EXIF: ${message}`); };
  abort();
  if (bytes.length < 2 || bytes[0] !== 0xff || bytes[1] !== 0xd8) fail("missing SOI");
  let offset = 2, segments = 0, exif = false;
  let orientation: JpegExifOrientationInspection["orientation"];
  const finish = (stoppedAt: "scan" | "end"): JpegExifOrientationInspection => ({
    status: !exif ? "no-exif" : orientation === undefined ? "no-orientation" : "orientation",
    ...(orientation === undefined ? {} : { orientation }), stoppedAt,
  });
  while (offset < bytes.length) {
    abort(); if (++segments > maxSegments) throw new RangeError("JPEG maxSegments exceeded.");
    if (bytes[offset++] !== 0xff) fail("expected marker");
    while (offset < bytes.length && bytes[offset] === 0xff) offset++;
    if (offset >= bytes.length) fail("truncated marker");
    const marker = bytes[offset++]!;
    if (marker === 0xd9) return finish("end");
    if (marker === 0x00 || marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) fail("unexpected standalone marker before scan");
    if (offset + 2 > bytes.length) fail("truncated segment length");
    const length = (bytes[offset]! << 8) | bytes[offset + 1]!;
    if (length < 2 || offset + length > bytes.length) fail("invalid segment length");
    const start = offset + 2, end = offset + length;
    const exifPrefix = marker === 0xe1 && end - start >= 4 && bytes[start] === 0x45 && bytes[start+1] === 0x78 && bytes[start+2] === 0x69 && bytes[start+3] === 0x66;
    if (exifPrefix) {
      if (end - start < 6 || bytes[start+4] !== 0 || bytes[start+5] !== 0) fail("malformed EXIF identifier");
      if (exif) fail("multiple EXIF segments");
      exif = true;
      const base = start + 6;
      const check = (at: number, length: number) => { if (at < base || at + length > end) fail("TIFF offset outside APP1"); };
      check(base, 8);
      const little = bytes[base] === 0x49 && bytes[base+1] === 0x49;
      if (!little && !(bytes[base] === 0x4d && bytes[base+1] === 0x4d)) fail("TIFF byte order");
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      const u16 = (at: number) => { check(at, 2); return view.getUint16(at, little); };
      const u32 = (at: number) => { check(at, 4); return view.getUint32(at, little); };
      if (u16(base + 2) !== 42) fail("TIFF magic");
      const ifdOffset = u32(base + 4); if (ifdOffset < 8) fail("invalid IFD0 offset");
      const ifd = base + ifdOffset, count = u16(ifd);
      if (count > maxIfdEntries) throw new RangeError("JPEG maxIfdEntries exceeded.");
      check(ifd + 2, count * 12 + 4);
      let found = false;
      for (let index = 0; index < count; index++) {
        abort(); const entry = ifd + 2 + index * 12;
        if (u16(entry) !== 0x0112) continue;
        if (found) fail("duplicate orientation"); found = true;
        if (u16(entry + 2) !== 3 || u32(entry + 4) !== 1) fail("invalid orientation type or count");
        const value = u16(entry + 8); if (value < 1 || value > 8) fail("invalid orientation value");
        orientation = value as JpegExifOrientationInspection["orientation"];
      }
    }
    offset = end;
    if (marker === 0xda) return finish("scan");
  }
  return fail("missing scan or EOI");
}
