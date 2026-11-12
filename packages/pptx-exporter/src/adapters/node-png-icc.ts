import { inflateSync } from "node:zlib";
import { capturePngInput } from "../assets/inspect-png.js";
import { capturedPngIccCompressed, inspectCapturedPngIcc, validatePngIccOptions, type PngIccInspection, type PngIccInspectionOptions } from "../assets/inspect-png-icc.js";

export interface ExtractPngIccOptions extends PngIccInspectionOptions {
  maxProfileBytes?: number;
}
export interface ExtractedPngIcc {
  inspection: PngIccInspection;
  /** Detached bytes when an iCCP chunk exists; the ICC contents are not validated. */
  profileBytes?: Uint8Array;
}

/** Node-only bounded extraction; PNG envelope and zlib stream are checked, not ICC semantics. */
export function extractPngIccProfile(input: Uint8Array, options: ExtractPngIccOptions = {}): ExtractedPngIcc {
  const { maxBytes, maxChunks, maxCompressedProfileBytes, signal, maxProfileBytes } = options;
  const captured = { ...(maxBytes === undefined ? {} : { maxBytes }), ...(maxChunks === undefined ? {} : { maxChunks }), ...(maxCompressedProfileBytes === undefined ? {} : { maxCompressedProfileBytes }), ...(signal === undefined ? {} : { signal }) };
  const limit = maxProfileBytes ?? 16 * 1024 * 1024;
  if (!Number.isSafeInteger(limit) || limit <= 0) throw new RangeError("maxProfileBytes must be a positive safe integer.");
  validatePngIccOptions(captured);
  const bytes = capturePngInput(input, captured), inspection = inspectCapturedPngIcc(bytes, captured);
  if (!inspection.profile) return { inspection };
  const compressed = capturedPngIccCompressed(bytes);
  signal?.throwIfAborted();
  try {
    // Node returns { buffer, engine } for info:true; the installed Node declarations omit this overload.
    const result = inflateSync(compressed, { info: true, maxOutputLength: limit }) as unknown as { buffer: Uint8Array; engine: { bytesWritten: number } };
    if (result.engine.bytesWritten !== compressed.length) throw new Error("Trailing data after PNG ICC zlib stream.");
    signal?.throwIfAborted();
    return { inspection, profileBytes: new Uint8Array(result.buffer) };
  } catch (error) {
    signal?.throwIfAborted();
    if ((error as NodeJS.ErrnoException)?.code === "ERR_BUFFER_TOO_LARGE") throw new RangeError("PNG maxProfileBytes exceeded.");
    throw error;
  }
}
