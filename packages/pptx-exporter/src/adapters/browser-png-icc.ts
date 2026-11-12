import { capturePngInput } from "../assets/inspect-png.js";
import { capturedPngIccCompressed, inspectCapturedPngIcc, validatePngIccOptions, type PngIccInspection, type PngIccInspectionOptions } from "../assets/inspect-png-icc.js";
import { withAbort } from "../assets/limits.js";

export interface BrowserExtractPngIccOptions extends PngIccInspectionOptions { maxProfileBytes?: number }
export interface BrowserExtractedPngIcc { inspection: PngIccInspection; profileBytes?: Uint8Array }

/** Explicit browser zlib extraction; neither ICC semantics nor browser-internal allocations are validated. */
export async function extractPngIccProfileInBrowser(input: Uint8Array, options: BrowserExtractPngIccOptions = {}): Promise<BrowserExtractedPngIcc> {
  const { maxBytes, maxChunks, maxCompressedProfileBytes, signal, maxProfileBytes } = options;
  const captured = { ...(maxBytes === undefined ? {} : { maxBytes }), ...(maxChunks === undefined ? {} : { maxChunks }), ...(maxCompressedProfileBytes === undefined ? {} : { maxCompressedProfileBytes }), ...(signal === undefined ? {} : { signal }) };
  const limit = maxProfileBytes ?? 16 * 1024 * 1024;
  if (!Number.isSafeInteger(limit) || limit <= 0) throw new RangeError("maxProfileBytes must be a positive safe integer.");
  validatePngIccOptions(captured);
  const bytes = capturePngInput(input, captured), inspection = inspectCapturedPngIcc(bytes, captured);
  if (!inspection.profile) return { inspection };
  if (typeof DecompressionStream !== "function") throw new Error("DecompressionStream is required to extract PNG ICC profiles.");
  signal?.throwIfAborted();
  const compressed = new Uint8Array(capturedPngIccCompressed(bytes));
  const reader = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate")).getReader();
  const chunks: Uint8Array[] = []; let total = 0;
  try {
    while (true) {
      const result = await withAbort(reader.read(), signal); signal?.throwIfAborted();
      if (result.done) break;
      total += result.value.length;
      if (total > limit) throw new RangeError("PNG maxProfileBytes exceeded.");
      chunks.push(new Uint8Array(result.value));
    }
    const profileBytes = new Uint8Array(total); let offset = 0;
    for (const chunk of chunks) { profileBytes.set(chunk, offset); offset += chunk.length; }
    return { inspection, profileBytes };
  } finally {
    void reader.cancel().catch(() => {}); reader.releaseLock();
  }
}
