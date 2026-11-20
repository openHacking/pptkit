import { captureImageInput } from "../images/capture-image-input.js";

export interface IsoBmffInspectionOptions { maxBytes?: number; maxBoxes?: number; signal?: AbortSignal }
export interface IsoBmffBox { readonly type: string; readonly offset: number; readonly size: number; readonly headerSize: 8 | 16 }
export interface IsoBmffInspection {
  readonly majorBrand: string;
  readonly minorVersion: number;
  readonly compatibleBrands: readonly string[];
  readonly boxes: readonly IsoBmffBox[];
}

/** Inspect only top-level ISO BMFF box framing and its file-type declaration. */
export function inspectIsoBmff(input: Uint8Array, options: IsoBmffInspectionOptions = {}): IsoBmffInspection {
  const maxBytes = options.maxBytes ?? 64 * 1024 * 1024, maxBoxes = options.maxBoxes ?? 4096;
  if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0 || maxBytes > 64 * 1024 * 1024) throw new RangeError("maxBytes must be between 1 and 67108864.");
  if (!Number.isSafeInteger(maxBoxes) || maxBoxes <= 0 || maxBoxes > 4096) throw new RangeError("maxBoxes must be between 1 and 4096.");
  const bytes = captureImageInput(input, "ISO BMFF", maxBytes, options.signal);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tag = (offset: number) => String.fromCharCode(...bytes.subarray(offset, offset + 4));
  const invalid = (reason: string): never => { throw new TypeError(`Invalid or unsupported ISO BMFF framing: ${reason}.`); };
  const boxes: IsoBmffBox[] = [];
  let at = 0, majorBrand: string | undefined, minorVersion = 0, compatibleBrands: string[] = [];
  while (at < bytes.length) {
    if (options.signal?.aborted) throw options.signal.reason ?? new Error("ISO BMFF inspection aborted.");
    if (boxes.length >= maxBoxes) throw new RangeError("ISO BMFF maxBoxes exceeded.");
    if (bytes.length - at < 8) return invalid("truncated box header");
    const shortSize = view.getUint32(at), type = tag(at + 4);
    let size = shortSize, headerSize: 8 | 16 = 8;
    if (shortSize === 0) size = bytes.length - at;
    else if (shortSize === 1) {
      if (bytes.length - at < 16) return invalid("truncated extended box header");
      headerSize = 16;
      const high = view.getUint32(at + 8), low = view.getUint32(at + 12);
      if (high !== 0) return invalid("extended box exceeds input");
      size = low;
    }
    if (size < headerSize || size > bytes.length - at) return invalid("invalid or truncated box size");
    if (type === "ftyp") {
      if (majorBrand !== undefined || size - headerSize < 8 || (size - headerSize - 8) % 4 !== 0) return invalid("invalid or duplicate file-type box");
      const payload = at + headerSize, brandCount = (size - headerSize - 8) / 4;
      if (brandCount > 4096) throw new RangeError("ISO BMFF compatible brand limit exceeded.");
      majorBrand = tag(payload); minorVersion = view.getUint32(payload + 4);
      compatibleBrands = Array.from({ length: brandCount }, (_, index) => tag(payload + 8 + index * 4));
    }
    boxes.push(Object.freeze({ type, offset: at, size, headerSize }));
    at += size;
  }
  if (majorBrand === undefined) return invalid("missing file-type box");
  return Object.freeze({ majorBrand, minorVersion, compatibleBrands: Object.freeze(compatibleBrands), boxes: Object.freeze(boxes) });
}
