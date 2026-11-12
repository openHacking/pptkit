import { captureImageInput } from "./capture-image-input.js";

export interface IccProfileInspectionOptions { maxBytes?: number; maxTags?: number; signal?: AbortSignal }
export interface IccProfileTag { signature: string; offset: number; size: number }
export interface IccProfileInspection {
  byteLength: number;
  versionMajor: number;
  profileClass: string;
  colorSpace: string;
  connectionSpace: string;
  tags: IccProfileTag[];
}

/** Bounded header/directory inventory; tag content and profile semantics are not validated. */
export function inspectIccProfile(input: Uint8Array, options: IccProfileInspectionOptions = {}): IccProfileInspection {
  const { maxBytes = 16 * 1024 * 1024, maxTags = 4096, signal } = options;
  if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0) throw new RangeError("maxBytes must be a positive safe integer.");
  if (!Number.isSafeInteger(maxTags) || maxTags <= 0) throw new RangeError("maxTags must be a positive safe integer.");
  const bytes = captureImageInput(input, "ICC", maxBytes, signal);
  const abort = () => { if (signal?.aborted) throw signal.reason ?? new Error("ICC inspection aborted."); };
  const fail = (reason: string): never => { throw new Error(`Invalid ICC profile directory: ${reason}`); };
  abort(); if (bytes.length < 132) fail("truncated header or tag count");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint32(0) !== bytes.length) fail("declared profile size does not match input");
  if (view.getUint32(36) !== 0x61637370) fail("missing acsp signature");
  const count = view.getUint32(128);
  if (count > maxTags) throw new RangeError("ICC maxTags exceeded.");
  const tableEnd = 132 + count * 12;
  if (tableEnd > bytes.length) fail("truncated tag table");
  const signature = (at: number) => String.fromCharCode(bytes[at]!, bytes[at + 1]!, bytes[at + 2]!, bytes[at + 3]!);
  const tags: IccProfileTag[] = [], signatures = new Set<string>();
  for (let i = 0; i < count; i++) {
    abort(); const at = 132 + i * 12, name = signature(at), offset = view.getUint32(at + 4), size = view.getUint32(at + 8);
    if (signatures.has(name)) fail("duplicate tag signature"); signatures.add(name);
    if (offset % 4 !== 0 || offset < tableEnd || size === 0 || offset > bytes.length || size > bytes.length - offset) fail("tag extent or alignment");
    tags.push({ signature: name, offset, size });
  }
  const ranges = [...tags].sort((a, b) => a.offset - b.offset || a.size - b.size);
  for (let i = 1; i < ranges.length; i++) {
    abort(); const previous = ranges[i - 1]!, current = ranges[i]!;
    if (current.offset < previous.offset + previous.size && !(current.offset === previous.offset && current.size === previous.size)) fail("partially overlapping tag data");
  }
  return { byteLength: bytes.length, versionMajor: bytes[8]!, profileClass: signature(12), colorSpace: signature(16), connectionSpace: signature(20), tags };
}
