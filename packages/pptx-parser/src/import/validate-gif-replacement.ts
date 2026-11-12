import { PptxImportError } from "../types/import.js";

/** Validate bounded GIF framing, LZW output and palette references, not playback timing. */
export async function validateGifReplacement(bytes: Uint8Array, maxWork: number, maxDecodedBytes: number, usage: { work: number; decodedBytes?: number }, checkpoint: () => Promise<void>): Promise<boolean> {
  const invalid = new Error("Invalid GIF structure or LZW stream."); let position = 0, blocks = 0, pending = 0;
  const fail = (): never => { throw invalid; };
  const charge = (count: number): Promise<void> | undefined => {
    usage.work += count; pending += count;
    if (usage.work > maxWork || blocks > 1000000) throw new PptxImportError("save-limit-exceeded", "GIF replacement inspection exceeds maxSaveValidationWork.");
    if (pending >= 32768) { pending = 0; return checkpoint(); }
  };
  const byte = () => { if (position >= bytes.length) return fail(); return bytes[position++]!; };
  const word = () => byte() | (byte() << 8);
  const skip = (length: number) => { if (length > bytes.length - position) fail(); position += length; };
  const subBlocks = async (retain: boolean) => {
    const spans: { start: number; length: number }[] = []; let size = 0;
    for (;;) {
      blocks++; const due = charge(1); if (due) await due; const length = byte(); if (!length) break;
      if (length > bytes.length - position) fail();
      if (retain) spans.push({ start: position, length }); size += length; skip(length); const dataDue = charge(length); if (dataDue) await dataDue;
    }
    if (!retain) return new Uint8Array();
    const data = new Uint8Array(size); let offset = 0;
    for (const span of spans) { data.set(bytes.subarray(span.start, span.start + span.length), offset); offset += span.length; }
    return data;
  };
  try {
    await checkpoint();
    if (bytes.length < 13 || !([71,73,70,56,55,97].every((v, i) => bytes[i] === v) || [71,73,70,56,57,97].every((v, i) => bytes[i] === v))) return false;
    position = 6; const width = word(), height = word(), flags = byte(); byte(); byte();
    if (!width || !height) return false;
    const globalColors = flags & 0x80 ? 1 << ((flags & 7) + 1) : 0;
    if (globalColors) skip(globalColors * 3);
    let frames = 0, control: { transparent?: number } | undefined;
    for (;;) {
      blocks++; const due = charge(1); if (due) await due; const marker = byte();
      if (marker === 0x3b) return frames > 0 && position === bytes.length && control === undefined;
      if (marker === 0x2c) {
        if (++frames > 10000) throw new PptxImportError("save-limit-exceeded", "GIF replacement exceeds its frame budget.");
        const left = word(), top = word(), frameWidth = word(), frameHeight = word(), frameFlags = byte();
        if (!frameWidth || !frameHeight || left + frameWidth > width || top + frameHeight > height || (frameFlags & 0x18)) return false;
        const localColors = frameFlags & 0x80 ? 1 << ((frameFlags & 7) + 1) : 0, colors = localColors || globalColors;
        if (!colors || (control?.transparent !== undefined && control.transparent >= colors)) return false;
        if (localColors) skip(localColors * 3);
        const minCodeSize = byte(); if (minCodeSize < 2 || minCodeSize > 8) return false;
        const compressed = await subBlocks(true), pixels = frameWidth * frameHeight;
        if (pixels > maxDecodedBytes - (usage.decodedBytes ?? 0)) throw new PptxImportError("save-limit-exceeded", "GIF decoded frames exceed maxSaveValidationBytes.");
        usage.decodedBytes = (usage.decodedBytes ?? 0) + pixels;
        if (!await validateLzw(compressed, minCodeSize, colors, pixels, charge)) return false;
        control = undefined; continue;
      }
      if (marker !== 0x21) return false;
      const label = byte();
      if (label === 0xf9) {
        if (control || byte() !== 4) return false;
        const controlFlags = byte(); word(); const transparent = byte();
        if (byte() !== 0 || (controlFlags & 0xe0) || ((controlFlags >>> 2) & 7) > 3) return false;
        control = controlFlags & 1 ? { transparent } : {};
      } else if (label === 0xff || label === 0x01) {
        const size = byte(); if (size !== (label === 0xff ? 11 : 12)) return false;
        skip(size); await subBlocks(false); if (label === 0x01) control = undefined;
      } else await subBlocks(false);
    }
  } catch (error) { if (error === invalid) return false; throw error; }
}

async function validateLzw(data: Uint8Array, minimum: number, colors: number, expected: number, charge: (count: number) => Promise<void> | undefined): Promise<boolean> {
  const clear = 1 << minimum, end = clear + 1, prefix = new Uint16Array(4096), suffix = new Uint8Array(4096), stack = new Uint8Array(4097);
  let bit = 0, width = minimum + 1, next = end + 1, previous = -1, previousFirst = 0, written = 0;
  const read = () => { if (bit + width > data.length * 8) return -1; let value = 0; for (let i = 0; i < width; i++, bit++) value |= ((data[bit >>> 3]! >>> (bit & 7)) & 1) << i; return value; };
  for (;;) {
    const due = charge(1); if (due) await due; const code = read(); if (code < 0) return false;
    if (code === clear) { width = minimum + 1; next = end + 1; previous = -1; continue; }
    if (code === end) return written === expected && data.length * 8 - bit < 8;
    if (code > next || (code === next && previous < 0) || code >= 4096) return false;
    const special = code === next; let cursor = special ? previous : code, count = 0;
    if (special) stack[count++] = previousFirst;
    while (cursor >= clear) {
      if (cursor < end + 1 || cursor >= next || count >= 4096) return false;
      stack[count++] = suffix[cursor]!; cursor = prefix[cursor]!;
    }
    if (cursor < 0 || count >= stack.length) return false;
    stack[count++] = cursor; const first = stack[count - 1]!;
    if (written + count > expected) return false;
    const pixelsDue = charge(count); if (pixelsDue) await pixelsDue; written += count;
    while (count) if (stack[--count]! >= colors) return false;
    if (previous >= 0 && next < 4096) {
      prefix[next] = previous; suffix[next] = first; next++;
      if (next === 1 << width && width < 12) width++;
    }
    previous = code; previousFirst = first;
  }
}
