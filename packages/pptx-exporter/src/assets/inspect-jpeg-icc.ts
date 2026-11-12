import { captureImageInput } from "./capture-image-input.js";

/** ICC chunk inventory before the first JPEG scan, not profile validation or color management. */
export interface JpegIccInspection {
  status: "no-profile" | "complete" | "incomplete";
  stoppedAt: "scan" | "end";
  chunkCount: number;
  receivedChunks: number;
  missingChunks: number[];
  profileBytes: number;
}
export interface JpegIccInspectionOptions {
  maxBytes?: number;
  maxSegments?: number;
  maxProfileBytes?: number;
  signal?: AbortSignal;
}
export interface ExtractedJpegIcc { inspection: JpegIccInspection; profileBytes?: Uint8Array }

export function inspectJpegIcc(input: Uint8Array, options: JpegIccInspectionOptions = {}): JpegIccInspection {
  return scanJpegIcc(input, options).inspection;
}

/** Reassemble complete pre-scan chunks; ICC contents and later JPEG scans are not validated. */
export function extractJpegIccProfile(input: Uint8Array, options: JpegIccInspectionOptions = {}): ExtractedJpegIcc {
  const { inspection, bytes, chunks, signal } = scanJpegIcc(input, options);
  if (inspection.status === "no-profile") return { inspection };
  if (inspection.status === "incomplete") throw new Error("Cannot extract an incomplete JPEG ICC profile.");
  const profileBytes = new Uint8Array(inspection.profileBytes); let offset = 0;
  for (let index = 1; index <= inspection.chunkCount; index++) {
    if (signal?.aborted) throw signal.reason ?? new Error("JPEG inspection aborted.");
    const range = chunks.get(index)!; profileBytes.set(bytes.subarray(range.start, range.end), offset); offset += range.end - range.start;
  }
  return { inspection, profileBytes };
}

function scanJpegIcc(input: Uint8Array, options: JpegIccInspectionOptions): { inspection: JpegIccInspection; bytes: Uint8Array; chunks: Map<number, { start: number; end: number }>; signal?: AbortSignal } {
  const { maxBytes: inputLimit, maxSegments: segmentLimit, maxProfileBytes: profileLimit, signal } = options;
  const limit = (value: number | undefined, fallback: number, name: string): number => {
    const result = value === undefined ? fallback : value;
    if (!Number.isSafeInteger(result) || result <= 0) throw new RangeError(`${name} must be a positive safe integer.`);
    return result;
  };
  const maxBytes = limit(inputLimit, 64 * 1024 * 1024, "maxBytes");
  const maxSegments = limit(segmentLimit, 100000, "maxSegments");
  const maxProfileBytes = limit(profileLimit, 16 * 1024 * 1024, "maxProfileBytes");
  const bytes = captureImageInput(input, "JPEG", maxBytes, signal);
  const abort = () => { if (signal?.aborted) throw signal.reason ?? new Error("JPEG inspection aborted."); };
  const fail = (reason: string): never => { throw new Error(`Invalid JPEG ICC: ${reason}`); };
  if (bytes.length < 2 || bytes[0] !== 0xff || bytes[1] !== 0xd8) fail("missing SOI");
  let offset = 2, segments = 0, count = 0, profileBytes = 0;
  const chunks = new Map<number, { start: number; end: number }>(), identifier = [73, 67, 67, 95, 80, 82, 79, 70, 73, 76, 69];
  const finish = (stoppedAt: "scan" | "end") => {
    const missingChunks = Array.from({ length: count }, (_, i) => i + 1).filter(index => !chunks.has(index));
    const inspection: JpegIccInspection = { status: count === 0 ? "no-profile" : missingChunks.length ? "incomplete" : "complete", stoppedAt, chunkCount: count, receivedChunks: chunks.size, missingChunks, profileBytes };
    return { inspection, bytes, chunks, ...(signal === undefined ? {} : { signal }) };
  };
  while (offset < bytes.length) {
    abort(); if (++segments > maxSegments) throw new RangeError("JPEG maxSegments exceeded.");
    if (bytes[offset++] !== 0xff) fail("expected marker");
    while (offset < bytes.length && bytes[offset] === 0xff) offset++;
    if (offset >= bytes.length) fail("truncated marker");
    const marker = bytes[offset++]!;
    if (marker === 0xd9) return finish("end");
    if (marker === 0x00 || marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) fail("unexpected standalone marker before scan");
    if (offset + 2 > bytes.length) fail("truncated segment length");
    const length = (bytes[offset]! << 8) | bytes[offset + 1]!;
    if (length < 2 || offset + length > bytes.length) fail("invalid segment length");
    const start = offset + 2, end = offset + length;
    if (marker === 0xe2 && end - start >= identifier.length && identifier.every((value, index) => bytes[start + index] === value)) {
      if (end - start < 14 || bytes[start + 11] !== 0) fail("truncated or malformed ICC identifier");
      const sequence = bytes[start + 12]!, total = bytes[start + 13]!;
      if (total === 0 || sequence === 0 || sequence > total) fail("invalid chunk sequence");
      if (count !== 0 && count !== total) fail("inconsistent chunk count");
      if (chunks.has(sequence)) fail("duplicate chunk sequence");
      count = total; chunks.set(sequence, { start: start + 14, end }); profileBytes += end - start - 14;
      if (profileBytes > maxProfileBytes) throw new RangeError("JPEG maxProfileBytes exceeded.");
    }
    offset = end;
    if (marker === 0xda) return finish("scan");
  }
  return fail("missing scan or EOI");
}
