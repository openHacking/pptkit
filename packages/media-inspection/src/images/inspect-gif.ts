import { captureImageInput } from "./capture-image-input.js";
import { createGifLzwValidator } from "./validate-gif-lzw.js";

/** Structural metadata only: no LZW decompression, image validation or playback. */
export interface GifFrameInfo {
  left: number;
  top: number;
  width: number;
  height: number;
  delayCentiseconds: number;
  disposal: number;
  /** Present when the graphic control requests user input before continuing. */
  userInputFlag?: true;
  /** Nonzero reserved high bits of the graphic-control packed byte. */
  reservedControlBits?: number;
  transparentIndex?: number;
  interlaced: boolean;
}
export interface GifInspection {
  version: "87a" | "89a";
  width: number;
  height: number;
  frames: GifFrameInfo[];
  /** Plain-text graphic blocks are counted but not decoded into frames. */
  plainTextBlocks?: number;
  /** Unrecognized application or extension blocks, whose effects are unknown. */
  unknownExtensionBlocks?: number;
  /** Netscape repeat count: 0 means infinite; absent means no loop extension. */
  loopCount?: number;
}
export interface GifInspectionOptions {
  maxBytes?: number;
  maxFrames?: number;
  maxBlocks?: number;
  signal?: AbortSignal;
}

export function inspectGif(input: Uint8Array, options: GifInspectionOptions = {}): GifInspection {
  return inspectGifInternal(input, options, false);
}

/** Internal strict export profile; checks the LZW stream without rendering. */
export function inspectGifImageData(input: Uint8Array, options: GifInspectionOptions): GifInspection {
  return inspectGifInternal(input, options, true);
}

interface StaticGifInspection extends GifInspection { danglingGraphicControl?: true; reservedImageBits?: true }
export function inspectGifForStatic(input: Uint8Array, options: GifInspectionOptions): StaticGifInspection {
  return inspectGifInternal(input, options, false, true);
}

function inspectGifInternal(input: Uint8Array, options: GifInspectionOptions, validateImageData: boolean, staticDiagnostics = false): StaticGifInspection {
  const limit = (value: number | undefined, fallback: number, name: string): number => {
    const result = value === undefined ? fallback : value;
    if (!Number.isSafeInteger(result) || result <= 0) throw new RangeError(`${name} must be a positive safe integer.`);
    return result;
  };
  const maxBytes = limit(options.maxBytes, 64 * 1024 * 1024, "maxBytes");
  const maxFrames = limit(options.maxFrames, 10000, "maxFrames");
  const maxBlocks = limit(options.maxBlocks, 1000000, "maxBlocks");
  const bytes = captureImageInput(input, "GIF", maxBytes, options.signal);
  let position = 0, blocks = 0, remainingIndices = 128 * 1024 * 1024;
  const abort = () => { if (options.signal?.aborted) throw options.signal.reason ?? new Error("GIF inspection aborted."); };
  const fail = (message: string): never => { throw new Error(`Invalid GIF: ${message}`); };
  const need = (count: number) => { if (position + count > bytes.length) fail("truncated input"); };
  const byte = () => { need(1); return bytes[position++]!; };
  const word = () => { const lo = byte(); return lo | byte() << 8; };
  const skip = (count: number) => { need(count); position += count; };
  const charge = () => { abort(); if (++blocks > maxBlocks) throw new RangeError("GIF maxBlocks exceeded."); };
  const subBlocks = (visit?: (start: number, size: number) => void) => {
    for (;;) { charge(); const size = byte(); if (size === 0) return; need(size); visit?.(position, size); skip(size); }
  };
  abort(); need(13);
  const signature = String.fromCharCode(...bytes.subarray(0, 6));
  if (signature !== "GIF87a" && signature !== "GIF89a") fail("signature");
  position = 6;
  const width = word(), height = word(), packed = byte();
  byte(); byte();
  if (!width || !height) fail("zero logical screen size");
  const globalColors = packed & 0x80 ? 1 << ((packed & 7) + 1) : 0;
  skip(3 * globalColors);
  const result: StaticGifInspection = { version: signature === "GIF87a" ? "87a" : "89a", width, height, frames: [] };
  let control: { delayCentiseconds: number; disposal: number; userInputFlag?: true; reservedControlBits?: number; transparentIndex?: number } | undefined;
  for (;;) {
    charge();
    const marker = byte();
    if (marker === 0x3b) {
      if (position !== bytes.length) fail("trailing bytes");
      if (!result.frames.length) fail("no image frames");
      if (staticDiagnostics && control !== undefined) result.danglingGraphicControl = true;
      return result;
    }
    if (marker === 0x2c) {
      if (result.frames.length >= maxFrames) throw new RangeError("GIF maxFrames exceeded.");
      const left = word(), top = word(), frameWidth = word(), frameHeight = word(), flags = byte();
      if (!frameWidth || !frameHeight || left + frameWidth > width || top + frameHeight > height) fail("invalid frame rectangle");
      const localColors = flags & 0x80 ? 1 << ((flags & 7) + 1) : 0;
      if (staticDiagnostics && ((flags & 0x18) !== 0 || ((flags & 0x20) !== 0 && !localColors))) result.reservedImageBits = true;
      skip(3 * localColors);
      const codeSize = byte(); if (codeSize < 2 || codeSize > 8) fail("invalid LZW minimum code size");
      if (validateImageData) {
        const colors = localColors || globalColors, pixels = frameWidth * frameHeight;
        if (!colors) fail("image has no local or global color table");
        if (control?.transparentIndex !== undefined && control.transparentIndex >= colors) fail("transparent index is out of range");
        if (pixels > remainingIndices) throw new RangeError("GIF decoded indices exceed the 128 MiB per-asset limit.");
        remainingIndices -= pixels;
        const lzw = createGifLzwValidator(codeSize, colors, pixels);
        subBlocks((start, size) => lzw.push(bytes, start, size));
        lzw.finish();
      } else subBlocks();
      result.frames.push({ left, top, width: frameWidth, height: frameHeight, delayCentiseconds: control?.delayCentiseconds ?? 0,
        disposal: control?.disposal ?? 0, ...(control?.userInputFlag ? { userInputFlag: true as const } : {}), ...(control?.reservedControlBits ? { reservedControlBits: control.reservedControlBits } : {}), ...(control?.transparentIndex === undefined ? {} : { transparentIndex: control.transparentIndex }), interlaced: !!(flags & 0x40) });
      control = undefined;
      continue;
    }
    if (marker !== 0x21) fail("unexpected block marker");
    const label = byte();
    if (label === 0xf9) {
      if (byte() !== 4) fail("graphic control size");
      const flags = byte(), delayCentiseconds = word(), transparentIndex = byte();
      if (byte() !== 0) fail("graphic control terminator");
      if (control !== undefined) fail("duplicate graphic control");
      control = { delayCentiseconds, disposal: (flags >> 2) & 7, ...((flags & 2) ? { userInputFlag: true as const } : {}), ...((flags & 0xe0) ? { reservedControlBits: flags & 0xe0 } : {}), ...((flags & 1) ? { transparentIndex } : {}) };
    } else if (label === 0xff) {
      const size = byte(); if (size !== 11) fail("application identifier size");
      need(size); const identifier = String.fromCharCode(...bytes.subarray(position, position + size)); skip(size);
      const known = identifier === "NETSCAPE2.0" || identifier === "ANIMEXTS1.0";
      if (!known) result.unknownExtensionBlocks = (result.unknownExtensionBlocks ?? 0) + 1;
      let first = true;
      subBlocks((start, count) => {
        if (!first && known) fail("unexpected loop extension data");
        if (first && known) {
          if (count !== 3 || bytes[start] !== 1 || result.loopCount !== undefined) fail("invalid or duplicate loop extension");
          result.loopCount = bytes[start + 1]! | bytes[start + 2]! << 8;
        }
        first = false;
      });
      if (first && known) fail("missing loop data");
    } else if (label === 0x01) {
      if (byte() !== 12) fail("plain text header size");
      skip(12); subBlocks(); control = undefined;
      result.plainTextBlocks = (result.plainTextBlocks ?? 0) + 1;
    } else {
      subBlocks();
      if (label !== 0xfe) result.unknownExtensionBlocks = (result.unknownExtensionBlocks ?? 0) + 1;
    }
  }
}
