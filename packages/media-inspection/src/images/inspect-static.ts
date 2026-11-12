import { validateJpegStructure } from "../jpeg/structure.js";
import { captureImageInput } from "./capture-image-input.js";
import { inspectApng } from "./inspect-apng.js";
import { inspectGifForStatic } from "./inspect-gif.js";

export interface StaticImageInspectionOptions { maxBytes?: number; signal?: AbortSignal }
export interface StaticImageInspection {
  format: "png" | "gif" | "jpeg" | "unknown";
  status: "static" | "animation-declared" | "unverified";
}
const safePngChunks = new Set(["IHDR", "PLTE", "IDAT", "IEND", "tRNS", "gAMA", "cHRM", "sRGB", "iCCP", "sBIT", "bKGD", "hIST", "pHYs", "tEXt", "zTXt", "iTXt", "tIME", "eXIf", "sPLT", "cICP", "mDCv", "cLLi"]);

/** Conservative classification of named, bounded static image structures. */
export function inspectStaticImage(input: Uint8Array, options: StaticImageInspectionOptions = {}): StaticImageInspection {
  const maxBytes = options.maxBytes === undefined ? 64 * 1024 * 1024 : options.maxBytes;
  if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0 || maxBytes > 64 * 1024 * 1024) throw new RangeError("maxBytes must be between 1 and 67108864.");
  const bytes = captureImageInput(input, "Image", maxBytes, options.signal);
  const common = { maxBytes, ...(options.signal === undefined ? {} : { signal: options.signal }) };
  if (bytes.length >= 8 && [137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value)) {
    const result = inspectApng(bytes, common);
    if (result.png.unknownCriticalChunks.length) return { format: "png", status: "unverified" };
    if (result.animation !== undefined) return { format: "png", status: "animation-declared" };
    // A private ancillary extension can have consumer-specific behavior.
    for (let at = 8; at < bytes.length;) {
      const length = new DataView(bytes.buffer, bytes.byteOffset + at, 4).getUint32(0);
      const type = String.fromCharCode(...bytes.subarray(at + 4, at + 8));
      if (!safePngChunks.has(type)) return { format: "png", status: "unverified" };
      at += length + 12;
    }
    return { format: "png", status: "static" };
  }
  if (bytes.length >= 6 && String.fromCharCode(...bytes.subarray(0, 6)).match(/^GIF8[79]a$/)) {
    const result = inspectGifForStatic(bytes, common);
    if (result.frames.length > 1) return { format: "gif", status: "animation-declared" };
    const frame = result.frames[0]!;
    if (result.loopCount !== undefined || result.plainTextBlocks || result.unknownExtensionBlocks || result.danglingGraphicControl || result.reservedImageBits || frame.userInputFlag || frame.reservedControlBits || frame.disposal > 1) return { format: "gif", status: "unverified" };
    return { format: "gif", status: "static" };
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    validateJpegStructure(bytes, options.signal);
    return { format: "jpeg", status: "static" };
  }
  return { format: "unknown", status: "unverified" };
}
