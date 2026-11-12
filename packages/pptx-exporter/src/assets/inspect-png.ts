import { captureImageInput, inspectPng } from "@pptkit/media-inspection";
import type { PngInspection, PngInspectionOptions } from "@pptkit/media-inspection";

export { inspectPng };
export type { PngInspection, PngInspectionOptions };

/** Retains the exporter's captured-input contract for ICC adapters. */
export function capturePngInput(bytes: Uint8Array, options: PngInspectionOptions): Uint8Array {
  const maxBytes = options.maxBytes === undefined ? 64 * 1024 * 1024 : options.maxBytes;
  if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0) throw new RangeError("maxBytes must be a positive safe integer.");
  const maxChunks = options.maxChunks === undefined ? 100000 : options.maxChunks;
  if (!Number.isSafeInteger(maxChunks) || maxChunks <= 0) throw new RangeError("maxChunks must be a positive safe integer.");
  return captureImageInput(bytes, "PNG", maxBytes, options.signal);
}

/** Revalidates a detached captured input; media inspection takes its own snapshot. */
export function inspectCapturedPng(bytes: Uint8Array, options: PngInspectionOptions): PngInspection {
  return inspectPng(bytes, { ...options, maxBytes: Math.max(1, bytes.byteLength) });
}
