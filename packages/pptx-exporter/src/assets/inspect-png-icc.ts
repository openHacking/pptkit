import { capturePngInput, inspectCapturedPng, type PngInspection, type PngInspectionOptions } from "./inspect-png.js";

export interface PngIccInspectionOptions extends PngInspectionOptions {
  maxCompressedProfileBytes?: number;
}
export interface PngIccInspection {
  png: PngInspection;
  /** Reported coexistence only; this does not resolve color-chunk precedence. */
  hasSrgb: boolean;
  hasCicp: boolean;
  profile?: { name: string; compressedBytes: number };
}

/** Checks the iCCP envelope only; compressed data and ICC semantics are not decoded. */
export function inspectPngIcc(input: Uint8Array, options: PngIccInspectionOptions = {}): PngIccInspection {
  const { maxBytes, maxChunks, maxCompressedProfileBytes, signal } = options;
  const captured = { ...(maxBytes === undefined ? {} : { maxBytes }), ...(maxChunks === undefined ? {} : { maxChunks }), ...(maxCompressedProfileBytes === undefined ? {} : { maxCompressedProfileBytes }), ...(signal === undefined ? {} : { signal }) };
  validatePngIccOptions(captured);
  const bytes = capturePngInput(input, captured);
  return inspectCapturedPngIcc(bytes, captured);
}

export function validatePngIccOptions(options: PngIccInspectionOptions): void {
  const value = options.maxCompressedProfileBytes ?? 16 * 1024 * 1024;
  if (!Number.isSafeInteger(value) || value <= 0) throw new RangeError("maxCompressedProfileBytes must be a positive safe integer.");
}

/** Internal only: caller owns the captured input. */
export function inspectCapturedPngIcc(bytes: Uint8Array, options: PngIccInspectionOptions): PngIccInspection {
  const maxProfile = options.maxCompressedProfileBytes ?? 16 * 1024 * 1024;
  validatePngIccOptions(options);
  const png = inspectCapturedPng(bytes, options);
  const abort = () => { if (options.signal?.aborted) throw options.signal.reason ?? new Error("PNG inspection aborted."); };
  const fail = (reason: string): never => { throw new Error(`Invalid PNG ICC: ${reason}`); };
  const u32 = (at: number) => ((bytes[at]! * 0x1000000) + (bytes[at + 1]! << 16) + (bytes[at + 2]! << 8) + bytes[at + 3]!) >>> 0;
  let at = 8, seenPaletteOrData = false, hasSrgb = false, hasCicp = false;
  let profile: PngIccInspection["profile"];
  while (at < bytes.length) {
    abort(); const length = u32(at), data = at + 8, end = data + length;
    const name = String.fromCharCode(bytes[at + 4]!, bytes[at + 5]!, bytes[at + 6]!, bytes[at + 7]!);
    if (name === "PLTE" || name === "IDAT") seenPaletteOrData = true;
    if (name === "sRGB") hasSrgb = true;
    if (name === "cICP") hasCicp = true;
    if (name === "iCCP") {
      if (profile || seenPaletteOrData) fail("duplicate or late iCCP");
      let zero = data; while (zero < end && bytes[zero] !== 0 && zero - data <= 79) zero++;
      const nameLength = zero - data;
      if (nameLength < 1 || nameLength > 79 || zero >= end) fail("invalid profile name length or separator");
      let previousSpace = false;
      for (let i = data; i < zero; i++) {
        const value = bytes[i]!, space = value === 32;
        if (!(value >= 32 && value <= 126) && !(value >= 161 && value <= 255)) fail("invalid profile name character");
        if (space && (i === data || i === zero - 1 || previousSpace)) fail("invalid profile name spacing");
        previousSpace = space;
      }
      if (zero + 2 >= end || bytes[zero + 1] !== 0) fail("missing profile data or unsupported compression method");
      const compressedBytes = end - zero - 2;
      if (compressedBytes > maxProfile) throw new RangeError("PNG maxCompressedProfileBytes exceeded.");
      profile = { name: String.fromCharCode(...bytes.subarray(data, zero)), compressedBytes };
    }
    at = end + 4;
  }
  return { png, hasSrgb, hasCicp, ...(profile ? { profile } : {}) };
}

/** Call only after successful envelope inspection of the same captured input. */
export function capturedPngIccCompressed(bytes: Uint8Array): Uint8Array {
  let at = 8;
  while (at < bytes.length) {
    const length = new DataView(bytes.buffer, bytes.byteOffset + at, 4).getUint32(0);
    const data = at + 8, end = data + length;
    if (bytes[at + 4] === 105 && bytes[at + 5] === 67 && bytes[at + 6] === 67 && bytes[at + 7] === 80) {
      let separator = data; while (bytes[separator] !== 0) separator++;
      return bytes.subarray(separator + 2, end);
    }
    at = end + 4;
  }
  throw new Error("PNG ICC chunk disappeared from captured input.");
}
