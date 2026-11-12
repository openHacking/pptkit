import { Inflate } from "fflate";
import { PackageInspectionError, type ZipArchiveInspection, type ZipEntryInfo, type ZipInspectionLimits, type ZipInspectionOptions } from "../types/archive.js";
import { updateCrc32 } from "./crc32.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { validateDeflate } from "./validate-deflate.js";

export const DEFAULT_ZIP_LIMITS: Readonly<ZipInspectionLimits> = Object.freeze({ maxInputBytes: 64 * 1024 * 1024, maxEntries: 20000, maxEntryBytes: 64 * 1024 * 1024, maxExpandedBytes: 256 * 1024 * 1024, maxCompressionRatio: 1000, maxStalledInputBytes: 128 * 1024 });
export interface IndexedEntry { info: ZipEntryInfo; start: number; end: number; local: number; recordEnd: number; central: number; centralEnd: number }
function fail(code: PackageInspectionError["code"], message: string, entry?: string): never { throw new PackageInspectionError(code, message, entry); }
function limitsFor(options: ZipInspectionOptions): ZipInspectionLimits {
  if (!options || typeof options !== "object" || (options.limits !== undefined && (!options.limits || typeof options.limits !== "object" || Array.isArray(options.limits)))) throw new TypeError("Invalid ZIP inspection options");
  const limits = { ...DEFAULT_ZIP_LIMITS, ...options.limits };
  for (const [name, value] of Object.entries(limits)) if (!Number.isFinite(value) || value <= 0 || (name !== "maxCompressionRatio" && !Number.isSafeInteger(value))) throw new RangeError(`Invalid ZIP limit ${name}`);
  return limits;
}
function checkAbort(signal: AbortSignal | undefined): void { if (signal?.aborted) fail("operation-aborted", "Archive inspection was aborted."); }
function decodeName(bytes: Uint8Array, flags: number): string {
  if (!(flags & 0x800) && bytes.some((byte) => byte >= 128)) fail("archive-unsupported", "Legacy non-ASCII ZIP filename encoding is not supported.");
  try { return new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(bytes); }
  catch { return fail("archive-path-invalid", "ZIP filename is not valid UTF-8."); }
}
export function pathKey(name: string): string {
  let decoded: string;
  try { decoded = decodeURIComponent(name); } catch { return fail("archive-path-invalid", "Malformed percent encoding in package path.", name); }
  if (!decoded || /[\\\u0000-\u001f\u007f:#?]/.test(decoded) || decoded.startsWith("/") || decoded.includes("%")) fail("archive-path-invalid", "Unsafe or ambiguous package path.", name);
  const segments = (decoded.endsWith("/") ? decoded.slice(0, -1) : decoded).split("/");
  if (segments.some((segment) => !segment || segment === "." || segment === "..")) fail("archive-path-invalid", "Package paths cannot contain empty or traversal segments.", name);
  // This inspection profile rejects aliases instead of choosing one spelling.
  return segments.join("/").normalize("NFC").toLowerCase();
}
export function indexArchive(bytes: Uint8Array, limits: ZipInspectionLimits): IndexedEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const range = (start: number, length: number, end = bytes.length): void => { if (!Number.isSafeInteger(start) || start < 0 || length < 0 || start + length > end) fail("archive-invalid", "ZIP record extends outside its declared region."); };
  const u16 = (offset: number): number => { range(offset, 2); return view.getUint16(offset, true); };
  const u32 = (offset: number): number => { range(offset, 4); return view.getUint32(offset, true); };
  const candidates: number[] = [];
  for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 65557); offset--) if (u32(offset) === 0x06054b50 && offset + 22 + u16(offset + 20) === bytes.length) candidates.push(offset);
  if (candidates.length !== 1) fail("archive-invalid", "Expected one unambiguous terminal ZIP end record.");
  const end = candidates[0]!;
  if (u16(end + 4) || u16(end + 6) || u16(end + 8) !== u16(end + 10)) fail("archive-unsupported", "Split/multi-disk ZIP is not supported.");
  const count = u16(end + 10), directorySize = u32(end + 12), directory = u32(end + 16);
  if (count === 0xffff || directorySize === 0xffffffff || directory === 0xffffffff) fail("archive-unsupported", "ZIP64 is not supported by this inspection profile.");
  if (count > limits.maxEntries) fail("archive-limit-exceeded", "ZIP entry count exceeds maxEntries.");
  if (directory + directorySize !== end) fail("archive-invalid", "Central directory does not cover the archive tail.");
  const extra = (start: number, length: number): void => {
    const stop = start + length; range(start, length);
    while (start < stop) {
      range(start, 4, stop); const tag = u16(start), size = u16(start + 2); range(start + 4, size, stop);
      if ([0x0001, 0x7075, 0x9901].includes(tag)) fail("archive-unsupported", "ZIP64, alternate Unicode names and AES extra fields require unsupported interpretation.");
      start += 4 + size;
    }
  };
  const indexed: IndexedEntry[] = [], keys = new Map<string, boolean>();
  let cursor = directory, total = 0;
  for (let index = 0; index < count; index++) {
    range(cursor, 46, end);
    if (u32(cursor) !== 0x02014b50) fail("archive-invalid", "Invalid central directory signature.");
    const version = u16(cursor + 6), flags = u16(cursor + 8), method = u16(cursor + 10);
    if (version > 20 || (flags & ~0x80e) || (method !== 0 && method !== 8) || (method === 0 && (flags & 6))) fail("archive-unsupported", "Unsupported ZIP version, flags, encryption or compression method.");
    const crc32 = u32(cursor + 16), compressed = u32(cursor + 20), expanded = u32(cursor + 24);
    const nameLength = u16(cursor + 28), extraLength = u16(cursor + 30), commentLength = u16(cursor + 32), local = u32(cursor + 42);
    if (compressed === 0xffffffff || expanded === 0xffffffff || local === 0xffffffff || u16(cursor + 34)) fail("archive-unsupported", "ZIP64 or multi-disk entry is not supported.");
    range(cursor + 46, nameLength + extraLength + commentLength, end);
    const nameBytes = bytes.subarray(cursor + 46, cursor + 46 + nameLength), name = decodeName(nameBytes, flags), key = pathKey(name);
    if (keys.has(key)) fail("archive-duplicate-path", "Duplicate or canonically aliased archive path.", name);
    const directoryEntry = name.endsWith("/");
    keys.set(key, directoryEntry);
    // Unix and OS X creator systems both use Unix mode bits for file type.
    const host = u16(cursor + 4) >>> 8, unixHost = host === 3 || host === 19, unixType = (u32(cursor + 38) >>> 16) & 0xf000;
    if (unixHost && unixType !== 0 && unixType !== 0x8000 && unixType !== 0x4000) fail("archive-unsupported", "ZIP links and special filesystem objects are not supported.", name);
    if ((((u32(cursor + 38) & 16) !== 0 || unixHost && unixType === 0x4000) && !directoryEntry) || (unixHost && unixType === 0x8000 && directoryEntry)) fail("archive-path-invalid", "Directory attributes disagree with the entry name.", name);
    if (directoryEntry && expanded !== 0) fail("archive-invalid", "Directory entries cannot have expanded content.", name);
    total += expanded;
    if (expanded > limits.maxEntryBytes || total > limits.maxExpandedBytes || expanded / Math.max(1, compressed) > limits.maxCompressionRatio) fail("archive-limit-exceeded", "Declared expanded size or compression ratio exceeds budget.", name);
    extra(cursor + 46 + nameLength, extraLength);
    range(local, 30, directory);
    if (u32(local) !== 0x04034b50 || u16(local + 4) !== version || u16(local + 6) !== flags || u16(local + 8) !== method) fail("archive-invalid", "Local and central headers disagree.", name);
    const localNameLength = u16(local + 26), localExtraLength = u16(local + 28);
    range(local + 30, localNameLength + localExtraLength, directory);
    const localName = bytes.subarray(local + 30, local + 30 + localNameLength);
    if (localNameLength !== nameLength || localName.some((value, i) => value !== nameBytes[i])) fail("archive-invalid", "Local and central filenames disagree.", name);
    extra(local + 30 + localNameLength, localExtraLength);
    for (const [offset, expected] of [[14, crc32], [18, compressed], [22, expanded]]) {
      const value = u32(local + offset!);
      if (value !== expected && (!(flags & 8) || value !== 0)) fail("archive-invalid", "Local and central sizes/CRC disagree.", name);
    }
    const start = local + 30 + localNameLength + localExtraLength, dataEnd = start + compressed;
    range(start, compressed, directory);
    let recordEnd = dataEnd;
    if (flags & 8) {
      const descriptor = (offset: number): boolean => offset + 12 <= directory && u32(offset) === crc32 && u32(offset + 4) === compressed && u32(offset + 8) === expanded;
      if (dataEnd + 4 <= directory && u32(dataEnd) === 0x08074b50 && descriptor(dataEnd + 4)) recordEnd += 16;
      else if (descriptor(dataEnd)) recordEnd += 12;
      else fail("archive-invalid", "Data descriptor does not match the central directory.", name);
    }
    indexed.push({ info: { name, directory: directoryEntry, method: method === 0 ? "stored" : "deflate", compressedBytes: compressed, expandedBytes: expanded, crc32 }, local, start, end: dataEnd, recordEnd, central: cursor, centralEnd: cursor + 46 + nameLength + extraLength + commentLength });
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  // A slash-terminated key sorts directly before its descendants, even with
  // siblings such as "a-". Avoid constructing every ancestor of deep paths.
  const paths = [...keys].map(([key, directory]) => ({ key, prefix: `${key}/`, directory })).sort((a, b) => a.prefix < b.prefix ? -1 : a.prefix > b.prefix ? 1 : 0);
  for (let i = 0; i + 1 < paths.length; i++) {
    const path = paths[i]!;
    if (!path.directory && paths[i + 1]!.prefix.startsWith(path.prefix)) fail("archive-path-invalid", "An archive file is also used as a parent directory.", path.key);
  }
  if (cursor !== end) fail("archive-invalid", "Central directory length/count mismatch.");
  let covered = 0;
  for (const entry of [...indexed].sort((a, b) => a.local - b.local)) {
    if (entry.local !== covered) fail("archive-invalid", "Overlapping, prefixed or unindexed local data is not supported.", entry.info.name);
    covered = entry.recordEnd;
  }
  if (covered !== directory) fail("archive-invalid", "Unindexed data precedes the central directory.");
  return indexed;
}

/** Internal payload reader. Public inspection returns only detached inventory metadata. */
export async function readArchive(input: Uint8Array, options: ZipInspectionOptions = {}, retain = false): Promise<{ inspection: ZipArchiveInspection; parts: Map<string, Uint8Array> }> {
  const limits = limitsFor(options);
  if (!(input instanceof Uint8Array)) throw new TypeError("ZIP input must be Uint8Array");
  checkAbort(options.signal);
  const view = byteViewInfo(input);
  if (view.shared) throw new TypeError("ZIP input must not use SharedArrayBuffer.");
  if (view.length > limits.maxInputBytes) fail("archive-limit-exceeded", "Compressed input exceeds maxInputBytes.");
  const signal = options.signal;
  const bytes = copyByteView(input);
  const indexed = indexArchive(bytes, limits);
  await Promise.resolve(); checkAbort(signal);
  let total = 0, pushes = 0;
  const parts = new Map<string, Uint8Array>();
  for (const entry of indexed) {
    if (entry.info.method === "deflate") {
      await validateDeflate(bytes.subarray(entry.start, entry.end), entry.info.name, (size) => {
        if (size > limits.maxEntryBytes || total + size > limits.maxExpandedBytes || size / Math.max(1, entry.info.compressedBytes) > limits.maxCompressionRatio) fail("archive-limit-exceeded", "DEFLATE expanded output exceeds budget.", entry.info.name);
        if (size > entry.info.expandedBytes) fail("archive-invalid", "DEFLATE expanded size exceeds its ZIP declaration.", entry.info.name);
      }, signal);
    }
    let count = 0, crc = 0xffffffff, stalled = 0;
    const chunks: Uint8Array[] = [];
    const accept = (chunk: Uint8Array): void => {
      count += chunk.length; total += chunk.length;
      if (count > limits.maxEntryBytes || total > limits.maxExpandedBytes || count / Math.max(1, entry.info.compressedBytes) > limits.maxCompressionRatio) fail("archive-limit-exceeded", "Actual expanded output exceeds budget.", entry.info.name);
      if (count > entry.info.expandedBytes) fail("archive-invalid", "Actual expanded size exceeds its ZIP declaration.", entry.info.name);
      crc = updateCrc32(crc, chunk);
      if (chunk.length) { stalled = 0; if (retain) chunks.push(chunk.slice()); }
    };
    try {
      const inflater = entry.info.method === "deflate" ? new Inflate((chunk) => accept(chunk)) : undefined;
      if (entry.start === entry.end && inflater) fail("archive-invalid", "Deflate payload cannot be empty.", entry.info.name);
      const chunkSize = Math.min(1024, limits.maxStalledInputBytes);
      for (let offset = entry.start; offset < entry.end; offset += chunkSize) {
        checkAbort(signal);
        const stop = Math.min(entry.end, offset + chunkSize), chunk = bytes.subarray(offset, stop);
        if (inflater) { stalled += chunk.length; if (stalled > limits.maxStalledInputBytes) fail("archive-limit-exceeded", "Deflate stream exceeded the stalled-input budget.", entry.info.name); inflater.push(chunk, stop === entry.end); }
        else accept(chunk);
        if (++pushes % 32 === 0) { await new Promise<void>((resolve) => setTimeout(resolve, 0)); checkAbort(signal); }
      }
    } catch (error) {
      if (error instanceof PackageInspectionError) throw error;
      fail("archive-invalid", "Invalid compressed entry payload.", entry.info.name);
    }
    if (count !== entry.info.expandedBytes) fail("archive-invalid", "Actual expanded size differs from the ZIP declaration.", entry.info.name);
    if (((crc ^ 0xffffffff) >>> 0) !== entry.info.crc32) fail("archive-crc-mismatch", "Entry CRC-32 verification failed.", entry.info.name);
    if (retain && !entry.info.directory) {
      const data = new Uint8Array(count); let offset = 0;
      for (const chunk of chunks) { data.set(chunk, offset); offset += chunk.length; }
      parts.set(entry.info.name, data);
    }
  }
  checkAbort(signal);
  return { inspection: { entries: indexed.map((entry) => ({ ...entry.info })), inputBytes: bytes.length, expandedBytes: total }, parts };
}
