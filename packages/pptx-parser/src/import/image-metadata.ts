/** Bounded header metadata only; no pixel decode, EXIF normalization or I/O. */
export async function rasterMetadata(bytes: Uint8Array, checkpoint: () => Promise<void>): Promise<{ mimeType: string; width: number; height: number } | undefined> {
  const starts = (prefix: number[]) => prefix.every((byte, i) => bytes[i] === byte);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (starts([137, 80, 78, 71, 13, 10, 26, 10])) {
    if (bytes.length < 33 || view.getUint32(8) !== 13 || view.getUint32(12) !== 0x49484452) return undefined;
    const width = view.getUint32(16), height = view.getUint32(20);
    return width > 0 && height > 0 && width <= 2147483647 && height <= 2147483647 ? { mimeType: "image/png", width, height } : undefined;
  }
  if (starts([71, 73, 70, 56, 55, 97]) || starts([71, 73, 70, 56, 57, 97])) {
    if (bytes.length < 13) return undefined;
    const width = view.getUint16(6, true), height = view.getUint16(8, true);
    return width > 0 && height > 0 ? { mimeType: "image/gif", width, height } : undefined;
  }
  if (!starts([255, 216, 255])) return undefined;
  // JPEG APP/COM/table payloads are skipped by their bounded segment lengths;
  // a byte pattern inside them must never be mistaken for a frame marker.
  let cursor = 2;
  while (cursor < bytes.length) {
    await checkpoint();
    if (bytes[cursor++] !== 255) return undefined;
    while (bytes[cursor] === 255) {
      if (cursor % 65536 === 0) await checkpoint();
      cursor++;
    }
    const marker = bytes[cursor++];
    if (marker === undefined || marker === 0 || marker === 0xd8 || marker === 0xd9 || marker === 0xda) return undefined;
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (cursor + 2 > bytes.length) return undefined;
    const length = view.getUint16(cursor);
    if (length < 2 || cursor + length > bytes.length) return undefined;
    if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
      if (length < 8) return undefined;
      const height = view.getUint16(cursor + 3), width = view.getUint16(cursor + 5), components = bytes[cursor + 7]!;
      if (!width || !height || !components || length !== 8 + 3 * components) return undefined;
      return { mimeType: "image/jpeg", width, height };
    }
    cursor += length;
  }
  return undefined;
}
