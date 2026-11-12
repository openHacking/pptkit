/** Read only a bounded JPEG header prefix. This does not validate or decode the image. */
export function probeJpegPrecision(base64: string): number | 'incomplete' | undefined {
  // Decode at most 64 KiB of input, rounded to a complete Base64 quartet.
  const bytes = atob(base64.slice(0, 87380));
  const hasMore = base64.length > 87380;
  if (bytes.length < 4 || bytes.charCodeAt(0) !== 0xff || bytes.charCodeAt(1) !== 0xd8) return undefined;
  let position = 2;
  while (position < bytes.length) {
    if (bytes.charCodeAt(position++) !== 0xff) return undefined;
    while (position < bytes.length && bytes.charCodeAt(position) === 0xff) position++;
    if (position >= bytes.length) return hasMore ? 'incomplete' : undefined;
    const marker = bytes.charCodeAt(position++);
    if (marker === 0 || marker === 0xd8 || marker === 0xd9 || marker === 0xda) return undefined;
    if (marker === 0x01 || marker >= 0xd0 && marker <= 0xd7) continue;
    if (position + 2 > bytes.length) return hasMore ? 'incomplete' : undefined;
    const length = bytes.charCodeAt(position) * 256 + bytes.charCodeAt(position + 1);
    if (length < 2) return undefined;
    if (length > bytes.length - position) return hasMore ? 'incomplete' : undefined;
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return length >= 8 ? bytes.charCodeAt(position + 2) : undefined;
    }
    position += length;
  }
  return hasMore ? 'incomplete' : undefined;
}
