import { captureImageInput } from "./capture-image-input.js";

/** Bounded PNG chunk inventory; compressed pixels and ancillary semantics are not decoded. */
export interface PngInspection {
  width: number;
  height: number;
  bitDepth: number;
  colorType: number;
  interlaced: boolean;
  chunkCount: number;
  idatBytes: number;
  /** Animation-related chunks are present; their sequencing and playback are not validated. */
  animationChunks: boolean;
  /** Number of iCCP chunks; profile contents, placement and compression are not validated. */
  iccProfileChunks?: number;
  /** Unrecognized critical chunk types, in first-seen order. */
  unknownCriticalChunks: string[];
}
export interface PngInspectionOptions { maxBytes?: number; maxChunks?: number; signal?: AbortSignal }

const crcTable = Uint32Array.from({ length: 256 }, (_, n) => { let c = n; for (let i = 0; i < 8; i++) c = (c & 1) ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
const depths: Readonly<Record<number, readonly number[]>> = { 0: [1, 2, 4, 8, 16], 2: [8, 16], 3: [1, 2, 4, 8], 4: [8, 16], 6: [8, 16] };

/** Capture once so CRC and subsequent metadata reads see the same bytes. */
export function capturePngInput(bytes: Uint8Array, options: PngInspectionOptions): Uint8Array {
  const maxBytes = options.maxBytes === undefined ? 64 * 1024 * 1024 : options.maxBytes;
  if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0) throw new RangeError("maxBytes must be a positive safe integer.");
  const maxChunks = options.maxChunks === undefined ? 100000 : options.maxChunks;
  if (!Number.isSafeInteger(maxChunks) || maxChunks <= 0) throw new RangeError("maxChunks must be a positive safe integer.");
  return captureImageInput(bytes, "PNG", maxBytes, options.signal);
}

export function inspectPng(bytes: Uint8Array, options: PngInspectionOptions = {}): PngInspection {
  return inspectCapturedPng(capturePngInput(bytes, options), options);
}

/** Internal only: caller must own an immutable snapshot. */
export function inspectCapturedPng(bytes: Uint8Array, options: PngInspectionOptions): PngInspection {
  const maxChunks = options.maxChunks === undefined ? 100000 : options.maxChunks;
  if (!Number.isSafeInteger(maxChunks) || maxChunks <= 0) throw new RangeError("maxChunks must be a positive safe integer.");
  const abort = () => { if (options.signal?.aborted) throw options.signal.reason ?? new Error("PNG inspection aborted."); };
  const fail = (message: string): never => { throw new Error(`Invalid PNG: ${message}`); };
  const u32 = (i: number) => ((bytes[i]! * 0x1000000) + (bytes[i + 1]! << 16) + (bytes[i + 2]! << 8) + bytes[i + 3]!) >>> 0;
  abort(); if (bytes.length < 8 || ![137,80,78,71,13,10,26,10].every((v,i) => bytes[i] === v)) fail("signature");
  let position = 8, count = 0, width = 0, height = 0, bitDepth = 0, colorType = -1, interlaced = false, idatBytes = 0, seenIdat = false, afterIdat = false, palette = false, animationChunks = false, iccProfileChunks = 0;
  const unknownCriticalChunks: string[] = [];
  const seenCritical = new Set<string>();
  while (position < bytes.length) {
    abort(); if (++count > maxChunks) throw new RangeError("PNG maxChunks exceeded.");
    if (bytes.length - position < 12) fail("truncated chunk");
    const length = u32(position); if (length > 0x7fffffff || length > bytes.length - position - 12) fail("chunk length");
    const start = position + 4, data = position + 8, end = data + length;
    const code = String.fromCharCode(...bytes.subarray(start, data)); if (!/^[A-Za-z]{4}$/.test(code) || code[2] !== code[2]!.toUpperCase()) fail("chunk type");
    let crc = 0xffffffff; for (let i = start; i < end; i++) { if ((i & 65535) === 0) abort(); crc = crcTable[(crc ^ bytes[i]!) & 255]! ^ (crc >>> 8); }
    if (((crc ^ 0xffffffff) >>> 0) !== u32(end)) fail(`CRC for ${code}`);
    if (count === 1 && code !== "IHDR") fail("IHDR must be first");
    if (code === "IHDR") {
      if (count !== 1 || length !== 13) fail("IHDR size or order");
      width = u32(data); height = u32(data + 4); bitDepth = bytes[data + 8]!; colorType = bytes[data + 9]!;
      if (!width || width > 0x7fffffff || !height || height > 0x7fffffff || !depths[colorType]?.includes(bitDepth) || bytes[data + 10] !== 0 || bytes[data + 11] !== 0 || (bytes[data + 12] !== 0 && bytes[data + 12] !== 1)) fail("IHDR fields");
      interlaced = bytes[data + 12] === 1;
    } else if (code === "PLTE") {
      if (palette || seenIdat || colorType === 0 || colorType === 4 || !length || length > 768 || length % 3 !== 0 || (colorType === 3 && length / 3 > 2 ** bitDepth)) fail("PLTE size or order");
      palette = true;
    } else if (code === "IDAT") {
      if (afterIdat || (colorType === 3 && !palette)) fail("IDAT order or missing palette");
      seenIdat = true; idatBytes += length;
    } else if (code === "IEND") {
      if (!seenIdat || length !== 0 || end + 4 !== bytes.length) fail("IEND or trailing bytes");
      return { width, height, bitDepth, colorType, interlaced, chunkCount: count, idatBytes, animationChunks, ...(iccProfileChunks ? { iccProfileChunks } : {}), unknownCriticalChunks };
    } else {
      if (seenIdat) afterIdat = true;
      if (code === "iCCP") iccProfileChunks++;
      if (["acTL", "fcTL", "fdAT"].includes(code)) animationChunks = true;
      if (code[0] === code[0]!.toUpperCase() && !seenCritical.has(code)) { seenCritical.add(code); unknownCriticalChunks.push(code); }
    }
    position = end + 4;
  }
  return fail("missing IEND");
}
