import type { LoadedAsset } from "../types/internal.js";
import { captureImageInput } from "./capture-image-input.js";

/** APNG uses PNG signature/packaging; animation is not inferred from its MIME alone. */
export function packageImageMime(value: string): string { return value.trim().toLowerCase() === "image/apng" ? "image/png" : value; }

/** Header classification, not a full image decoder or fidelity certification. */
export function loadImageBytes(input: Uint8Array, declaredMime?: string, maxAssetBytes?: number): LoadedAsset {
  return classifyOwnedImageBytes(captureImageInput(input, "Asset", maxAssetBytes ?? Number.MAX_SAFE_INTEGER), declaredMime);
}

/** Only for an array exclusively owned by the loader. */
export function classifyOwnedImageBytes(data: Uint8Array, declaredMime?: string): LoadedAsset {
  const starts = (prefix: readonly number[]) => prefix.every((value, index) => data[index] === value);
  const mimeType = starts([137, 80, 78, 71, 13, 10, 26, 10]) ? "image/png"
    : starts([255, 216, 255]) ? "image/jpeg"
      : starts([71, 73, 70, 56, 55, 97]) || starts([71, 73, 70, 56, 57, 97]) ? "image/gif" : undefined;
  if (mimeType === undefined) throw new Error("Byte asset has an unsupported image signature; PNG, JPEG and GIF are supported.");
  if (declaredMime !== undefined && packageImageMime(declaredMime).trim().toLowerCase() !== mimeType) throw new Error(`Declared MIME ${declaredMime} does not match ${mimeType} signature.`);
  return { data, mimeType, extension: mimeType === "image/jpeg" ? ".jpg" : mimeType === "image/gif" ? ".gif" : ".png" };
}

export function decodeAssetBytes(value: string): Uint8Array {
  return Uint8Array.from(atob(value), (char) => char.charCodeAt(0));
}
