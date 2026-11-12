import { DEFAULT_ZIP_LIMITS, indexArchive, pathKey, readArchive } from "../archive/read-archive.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { updateCrc32 } from "../archive/crc32.js";
import { PackageInspectionError, type ZipInspectionOptions } from "../types/archive.js";

/** Rebase ZIP32 offsets while preserving untouched local/central records verbatim. */
export async function replaceArchiveEntries(
  input: Uint8Array, replacements: readonly { readonly name: string; readonly bytes: Uint8Array }[],
  options: ZipInspectionOptions & { readonly maxOutputBytes: number; readonly additions?: readonly { readonly name: string; readonly bytes: Uint8Array }[] },
): Promise<Uint8Array> {
  if (!(input instanceof Uint8Array) || !Array.isArray(replacements)) throw new TypeError("Invalid ZIP replacement input.");
  if (!Number.isSafeInteger(options.maxOutputBytes) || options.maxOutputBytes <= 0) throw new RangeError("Invalid ZIP replacement output limit.");
  const view = byteViewInfo(input);
  if (view.shared) throw new TypeError("ZIP replacement requires an ordinary byte buffer.");
  const signal = options.signal, limits = { ...DEFAULT_ZIP_LIMITS, ...options.limits }, maxOutput = options.maxOutputBytes, additions = options.additions ?? [];
  if (!Array.isArray(additions)) throw new TypeError("Invalid ZIP additions.");
  for (const [name, value] of Object.entries(limits)) if (!Number.isFinite(value) || value <= 0 || (name !== "maxCompressionRatio" && !Number.isSafeInteger(value))) throw new RangeError(`Invalid ZIP limit ${name}.`);
  const check = () => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "ZIP replacement was aborted."); };
  check();
  if (view.length > limits.maxInputBytes) throw new PackageInspectionError("archive-limit-exceeded", "ZIP replacement input exceeds maxInputBytes.");
  const bytes = copyByteView(input), patches = new Map<string, Uint8Array>();
  let patchBytes = 0;
  if (replacements.length > limits.maxEntries) throw new PackageInspectionError("archive-limit-exceeded", "ZIP replacements exceed maxEntries.");
  for (const replacement of replacements) {
    if (!replacement) throw new TypeError("Invalid ZIP replacement record.");
    const name = replacement.name, patch = replacement.bytes;
    if (typeof name !== "string" || !(patch instanceof Uint8Array)) throw new TypeError("Invalid ZIP replacement record.");
    if (patches.has(name)) throw new PackageInspectionError("archive-invalid", "Duplicate ZIP replacement.", name);
    const patchView = byteViewInfo(patch);
    if (patchView.shared) throw new TypeError("ZIP replacement requires ordinary byte buffers.");
    patchBytes += patchView.length;
    if (patchBytes > maxOutput || patchView.length > limits.maxEntryBytes) throw new PackageInspectionError("archive-limit-exceeded", "ZIP replacement payload exceeds output/entry budget.", name);
    patches.set(name, copyByteView(patch));
  }
  const newEntries: { name: string; nameBytes: Uint8Array; bytes: Uint8Array }[] = [], newNames = new Set<string>();
  if (additions.length > limits.maxEntries) throw new PackageInspectionError("archive-limit-exceeded", "ZIP additions exceed maxEntries.");
  for (const addition of additions) {
    if (!addition) throw new TypeError("Invalid ZIP addition record.");
    const name = addition.name, data = addition.bytes;
    if (typeof name !== "string" || !(data instanceof Uint8Array)) throw new TypeError("Invalid ZIP addition record.");
    if (!/^[A-Za-z0-9._/-]+$/.test(name) || name.startsWith("/") || name.endsWith("/") || name.split("/").some(part => !part || part === "." || part === "..")) throw new PackageInspectionError("archive-path-invalid", "ZIP addition requires a safe ASCII file path.", name);
    const key = name.toLowerCase(), nameBytes = new TextEncoder().encode(name);
    if (newNames.has(key)) throw new PackageInspectionError("archive-duplicate-path", "Duplicate ZIP addition.", name);
    if (nameBytes.length > 0xffff) throw new PackageInspectionError("archive-limit-exceeded", "ZIP addition name exceeds ZIP32 limit.", name);
    const item = byteViewInfo(data); if (item.shared) throw new TypeError("ZIP replacement requires ordinary byte buffers.");
    patchBytes += item.length;
    if (patchBytes > maxOutput || item.length > limits.maxEntryBytes) throw new PackageInspectionError("archive-limit-exceeded", "ZIP addition payload exceeds output/entry budget.", name);
    newNames.add(key); newEntries.push({ name, nameBytes, bytes: copyByteView(data) });
  }
  // Complete validation, including CRC/DEFLATE, precedes use of raw compressed records.
  await readArchive(bytes, { limits, ...(signal ? { signal } : {}) }); check();
  const indexed = indexArchive(bytes, limits), names = new Set(indexed.map(entry => entry.info.name));
  if (indexed.length + newEntries.length > limits.maxEntries || indexed.length + newEntries.length >= 0xffff) throw new PackageInspectionError("archive-limit-exceeded", "ZIP candidate exceeds entry count budget.");
  const paths = indexed.map(entry => ({ key: pathKey(entry.info.name), directory: entry.info.directory, name: entry.info.name }));
  for (const entry of newEntries) paths.push({ key: pathKey(entry.name), directory: false, name: entry.name });
  paths.sort((a, b) => `${a.key}/` < `${b.key}/` ? -1 : `${a.key}/` > `${b.key}/` ? 1 : 0);
  for (let i = 1; i < paths.length; i++) {
    const previous = paths[i - 1]!, current = paths[i]!;
    if (previous.key === current.key) throw new PackageInspectionError("archive-duplicate-path", "ZIP addition conflicts with an existing entry.", current.name);
    if (!previous.directory && current.key.startsWith(`${previous.key}/`)) throw new PackageInspectionError("archive-path-invalid", "A ZIP file cannot also be a parent directory.", previous.name);
  }
  for (const name of patches.keys()) if (!names.has(name)) throw new PackageInspectionError("archive-invalid", "ZIP replacement target is absent.", name);
  const segments: Uint8Array[] = [], offsets = new Map<string, number>(), headers = new Map<string, { crc: number; size: number }>();
  let size = 0;
  const add = (segment: Uint8Array) => {
    size += segment.length;
    if (size > maxOutput || size >= 0xffffffff) throw new PackageInspectionError("archive-limit-exceeded", "ZIP candidate exceeds output/ZIP32 budget.");
    segments.push(segment);
  };
  let expanded = 0, blocks = 0;
  for (const entry of [...indexed].sort((a, b) => a.local - b.local)) {
    check(); offsets.set(entry.info.name, size);
    const patch = patches.get(entry.info.name);
    expanded += patch?.length ?? entry.info.expandedBytes;
    if (expanded > limits.maxExpandedBytes) throw new PackageInspectionError("archive-limit-exceeded", "ZIP candidate exceeds maxExpandedBytes.");
    if (!patch) { add(bytes.subarray(entry.local, entry.recordEnd)); continue; }
    if (entry.info.directory) throw new PackageInspectionError("archive-invalid", "Directory payload replacement is forbidden.", entry.info.name);
    let crc = 0xffffffff;
    for (let offset = 0; offset < patch.length; offset += 65536) {
      check(); crc = updateCrc32(crc, patch.subarray(offset, offset + 65536));
      if (++blocks % 32 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
    }
    crc = (crc ^ 0xffffffff) >>> 0;
    const header = bytes.slice(entry.local, entry.start), view = new DataView(header.buffer);
    view.setUint16(6, view.getUint16(6, true) & ~0x0e, true); view.setUint16(8, 0, true);
    view.setUint32(14, crc, true); view.setUint32(18, patch.length, true); view.setUint32(22, patch.length, true);
    headers.set(entry.info.name, { crc, size: patch.length });
    add(header); add(patch);
  }
  const addedHeaders: { nameBytes: Uint8Array; offset: number; crc: number; size: number }[] = [];
  for (const entry of newEntries) {
    check(); expanded += entry.bytes.length;
    if (expanded > limits.maxExpandedBytes) throw new PackageInspectionError("archive-limit-exceeded", "ZIP candidate exceeds maxExpandedBytes.");
    let crc = 0xffffffff;
    for (let offset = 0; offset < entry.bytes.length; offset += 65536) {
      check(); crc = updateCrc32(crc, entry.bytes.subarray(offset, offset + 65536));
      if (++blocks % 32 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
    }
    crc = (crc ^ 0xffffffff) >>> 0;
    const local = new Uint8Array(30 + entry.nameBytes.length), view = new DataView(local.buffer);
    view.setUint32(0, 0x04034b50, true); view.setUint16(4, 20, true); view.setUint16(12, 0x21, true);
    view.setUint32(14, crc, true); view.setUint32(18, entry.bytes.length, true); view.setUint32(22, entry.bytes.length, true); view.setUint16(26, entry.nameBytes.length, true);
    local.set(entry.nameBytes, 30); addedHeaders.push({ nameBytes: entry.nameBytes, offset: size, crc, size: entry.bytes.length });
    add(local); add(entry.bytes);
  }
  const directory = size;
  for (const entry of indexed) {
    check();
    const central = bytes.slice(entry.central, entry.centralEnd), view = new DataView(central.buffer), header = headers.get(entry.info.name);
    view.setUint32(42, offsets.get(entry.info.name)!, true);
    if (header) {
      view.setUint16(8, view.getUint16(8, true) & ~0x0e, true); view.setUint16(10, 0, true);
      view.setUint32(16, header.crc, true); view.setUint32(20, header.size, true); view.setUint32(24, header.size, true);
    }
    add(central);
  }
  for (const entry of addedHeaders) {
    const central = new Uint8Array(46 + entry.nameBytes.length), view = new DataView(central.buffer);
    view.setUint32(0, 0x02014b50, true); view.setUint16(4, 20, true); view.setUint16(6, 20, true); view.setUint16(14, 0x21, true);
    view.setUint32(16, entry.crc, true); view.setUint32(20, entry.size, true); view.setUint32(24, entry.size, true); view.setUint16(28, entry.nameBytes.length, true); view.setUint32(42, entry.offset, true);
    central.set(entry.nameBytes, 46); add(central);
  }
  const directorySize = size - directory, end = indexed.length ? indexed[indexed.length - 1]!.centralEnd : 0;
  const tail = bytes.slice(end), tailView = new DataView(tail.buffer);
  tailView.setUint16(8, indexed.length + newEntries.length, true); tailView.setUint16(10, indexed.length + newEntries.length, true);
  tailView.setUint32(12, directorySize, true); tailView.setUint32(16, directory, true); add(tail);
  check();
  const output = new Uint8Array(size); let written = 0;
  for (const segment of segments) for (let start = 0; start < segment.length; start += 1024 * 1024) {
    check(); const slice = segment.subarray(start, start + 1024 * 1024); output.set(slice, written); written += slice.length;
    if (++blocks % 32 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
  }
  // Fresh header/size/CRC validation is independent of candidate construction.
  await readArchive(output, { limits: { ...limits, maxInputBytes: Math.max(limits.maxInputBytes, maxOutput) }, ...(signal ? { signal } : {}) });
  check(); return output;
}
