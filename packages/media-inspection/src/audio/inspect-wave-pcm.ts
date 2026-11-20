import { captureImageInput } from "../images/capture-image-input.js";

export interface WavePcmInspectionOptions { maxBytes?: number; maxChunks?: number; signal?: AbortSignal }
export interface WavePcmInspection {
  readonly format: "wave-pcm";
  readonly channels: 1 | 2;
  readonly sampleRate: number;
  readonly bitsPerSample: 8 | 16;
  readonly frameCount: number;
  readonly durationSeconds: number;
  readonly dataOffset: number;
  readonly dataLength: number;
}

/** Inspect a deliberately narrow RIFF/WAVE PCM profile without decoding samples. */
export function inspectWavePcm(input: Uint8Array, options: WavePcmInspectionOptions = {}): WavePcmInspection {
  const maxBytes = options.maxBytes ?? 64 * 1024 * 1024, maxChunks = options.maxChunks ?? 4096;
  if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0 || maxBytes > 64 * 1024 * 1024) throw new RangeError("maxBytes must be between 1 and 67108864.");
  if (!Number.isSafeInteger(maxChunks) || maxChunks <= 0 || maxChunks > 4096) throw new RangeError("maxChunks must be between 1 and 4096.");
  const bytes = captureImageInput(input, "WAVE", maxBytes, options.signal);
  const invalid = (reason: string): never => { throw new TypeError(`Invalid or unsupported WAVE PCM: ${reason}.`); };
  if (bytes.length < 12) return invalid("truncated RIFF header");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tag = (offset: number) => String.fromCharCode(...bytes.subarray(offset, offset + 4));
  if (tag(0) !== "RIFF" || tag(8) !== "WAVE") return invalid("RIFF/WAVE signature required");
  if (view.getUint32(4, true) + 8 !== bytes.length) return invalid("RIFF size mismatch");
  let at = 12, count = 0, fmt: { channels: number; sampleRate: number; bits: number; block: number } | undefined, data: { offset: number; length: number } | undefined;
  while (at < bytes.length) {
    if (options.signal?.aborted) throw options.signal.reason ?? new Error("WAVE inspection aborted.");
    if (++count > maxChunks) throw new RangeError("WAVE maxChunks exceeded.");
    if (bytes.length - at < 8) return invalid("truncated chunk header");
    const name = tag(at), length = view.getUint32(at + 4, true), start = at + 8, end = start + length, next = end + (length % 2);
    if (next > bytes.length) return invalid("truncated chunk or padding");
    if (name === "fmt ") {
      if (fmt || data || length !== 16) return invalid("expected one 16-byte fmt chunk before data");
      const format = view.getUint16(start, true), channels = view.getUint16(start + 2, true), sampleRate = view.getUint32(start + 4, true), rate = view.getUint32(start + 8, true), block = view.getUint16(start + 12, true), bits = view.getUint16(start + 14, true);
      if (format !== 1 || (channels !== 1 && channels !== 2) || (bits !== 8 && bits !== 16) || sampleRate === 0 || block !== channels * bits / 8 || rate !== sampleRate * block) return invalid("unsupported or inconsistent PCM format");
      fmt = { channels, sampleRate, bits, block };
    } else if (name === "data") {
      if (!fmt || data) return invalid("expected one data chunk after fmt");
      if (length % fmt.block !== 0) return invalid("data is not frame aligned");
      data = { offset: start, length };
    }
    at = next;
  }
  if (!fmt || !data) return invalid("missing fmt or data chunk");
  const frameCount = data.length / fmt.block;
  return Object.freeze({ format: "wave-pcm", channels: fmt.channels as 1 | 2, sampleRate: fmt.sampleRate, bitsPerSample: fmt.bits as 8 | 16, frameCount, durationSeconds: frameCount / fmt.sampleRate, dataOffset: data.offset, dataLength: data.length });
}
