import { updateCrc32 } from "../archive/crc32.js";
import { validateDeflate } from "../archive/validate-deflate.js";
import { Inflate } from "fflate";
import { PackageInspectionError } from "../types/archive.js";
import { PptxImportError } from "../types/import.js";

/** Validate container, scanline framing and indexed palette references. */
export async function validatePngReplacement(bytes: Uint8Array, maxWork: number, checkpoint: () => Promise<void>, usage: { work: number; decodedBytes?: number } = { work: 0 }, maxDecodedBytes = 128 * 1024 * 1024, signal?: AbortSignal): Promise<boolean> {
  if (bytes.length < 8 || ![137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value)) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const depths: Readonly<Record<number, readonly number[]>> = { 0: [1, 2, 4, 8, 16], 2: [8, 16], 3: [1, 2, 4, 8], 4: [8, 16], 6: [8, 16] };
  let position = 8, chunks = 0, width = 0, height = 0, colorType = -1, bitDepth = 0, interlaced = false, palette = false, paletteEntries = 0, transparency = false, seenData = false, afterData = false, dataBytes = 0;
  const dataSpans: { start: number; end: number }[] = [];
  while (position < bytes.length) {
    await checkpoint();
    if (++chunks > 100000 || ++usage.work > maxWork) throw new PptxImportError("save-limit-exceeded", "PNG replacement inspection exceeds its work budget.");
    if (bytes.length - position < 12) return false;
    const length = view.getUint32(position), start = position + 4, data = position + 8;
    if (length > 0x7fffffff || length > bytes.length - position - 12) return false;
    const end = data + length, type = String.fromCharCode(...bytes.subarray(start, data));
    if (!/^[A-Za-z]{4}$/.test(type) || type[2] !== type[2]!.toUpperCase()) return false;
    let crc = 0xffffffff;
    for (let offset = start; offset < end; offset += 65536) {
      const block = bytes.subarray(offset, Math.min(end, offset + 65536)); usage.work += block.length;
      if (usage.work > maxWork) throw new PptxImportError("save-limit-exceeded", "PNG replacement inspection exceeds its work budget.");
      crc = updateCrc32(crc, block); await checkpoint();
    }
    if (((crc ^ 0xffffffff) >>> 0) !== view.getUint32(end)) return false;
    if (chunks === 1 && type !== "IHDR") return false;
    if (type === "IHDR") {
      if (chunks !== 1 || length !== 13) return false;
      width = view.getUint32(data); height = view.getUint32(data + 4); bitDepth = bytes[data + 8]!; colorType = bytes[data + 9]!; interlaced = bytes[data + 12] === 1;
      if (!width || width > 0x7fffffff || !height || height > 0x7fffffff || !depths[colorType]?.includes(bitDepth) || bytes[data + 10] !== 0 || bytes[data + 11] !== 0 || (bytes[data + 12] !== 0 && bytes[data + 12] !== 1)) return false;
    } else if (type === "PLTE") {
      if (palette || transparency || seenData || colorType === 0 || colorType === 4 || !length || length > 768 || length % 3 !== 0 || (colorType === 3 && length / 3 > 2 ** bitDepth)) return false;
      palette = true; paletteEntries = length / 3;
    } else if (type === "tRNS") {
      if (transparency || seenData) return false;
      transparency = true;
      if (colorType === 3) {
        if (!palette || !length || length > paletteEntries) return false;
      } else if (colorType === 0 || colorType === 2) {
        const samples = colorType === 0 ? 1 : 3;
        if (length !== samples * 2) return false;
        for (let i = 0; i < samples; i++) if (view.getUint16(data + i * 2) >= 2 ** bitDepth) return false;
      } else return false;
    } else if (type === "IDAT") {
      if (afterData || (colorType === 3 && !palette)) return false;
      seenData = true; dataBytes += length; dataSpans.push({ start: data, end });
    } else if (type === "IEND") {
      if (!seenData || dataBytes < 6 || length !== 0 || end + 4 !== bytes.length) return false;
      const compressed = new Uint8Array(dataBytes); let offset = 0;
      for (const span of dataSpans) { compressed.set(bytes.subarray(span.start, span.end), offset); offset += span.end - span.start; await checkpoint(); }
      return validatePngData(compressed, width, height, bitDepth, colorType, interlaced, paletteEntries, maxWork, maxDecodedBytes, usage, checkpoint, signal);
    } else {
      if (seenData) afterData = true;
      if (type[0] === type[0]!.toUpperCase()) return false;
    }
    position = end + 4;
  }
  return false;
}

async function validatePngData(data: Uint8Array, width: number, height: number, bitDepth: number, colorType: number, interlaced: boolean, paletteEntries: number, maxWork: number, maxDecodedBytes: number, usage: { work: number; decodedBytes?: number }, checkpoint: () => Promise<void>, signal?: AbortSignal): Promise<boolean> {
  const cmf = data[0]!, flags = data[1]!;
  if ((cmf & 15) !== 8 || (cmf >>> 4) > 7 || ((cmf << 8) + flags) % 31 !== 0 || (flags & 32) !== 0) return false;
  const channels = ({ 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 } as Readonly<Record<number, number>>)[colorType]!;
  const geometry = interlaced ? [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]] : [[0, 0, 1, 1]];
  const passes: { rowBytes: number; rows: number; width: number }[] = []; let expected = 0, pixels = 0;
  for (const [x, y, dx, dy] of geometry) {
    const w = width <= x! ? 0 : Math.ceil((width - x!) / dx!), h = height <= y! ? 0 : Math.ceil((height - y!) / dy!);
    if (!w || !h) continue;
    const rowBytes = 1 + Math.ceil(w * channels * bitDepth / 8);
    if (rowBytes > Math.floor((maxDecodedBytes - (usage.decodedBytes ?? 0) - expected) / h)) throw new PptxImportError("save-limit-exceeded", "PNG decoded scanlines exceed maxSaveValidationBytes.");
    expected += rowBytes * h; pixels += w * h; passes.push({ rowBytes, rows: h, width: w });
  }
  usage.decodedBytes = (usage.decodedBytes ?? 0) + expected;
  const raw = data.subarray(2, data.length - 4), size = raw.length + expected;
  usage.work += 2 * size + (colorType === 3 ? expected + pixels : 0);
  if (usage.work > maxWork) throw new PptxImportError("save-limit-exceeded", "PNG data inspection exceeds maxSaveValidationWork.");
  const invalid = new Error("Invalid PNG scanline stream."); let preflight = 0;
  try { await validateDeflate(raw, "PNG IDAT", value => { if (value > expected) throw invalid; preflight = value; }, signal, 1 << ((cmf >>> 4) + 8)); }
  catch (error) { if (error instanceof PackageInspectionError && error.code === "operation-aborted") throw error; if (error instanceof PptxImportError) throw error; return false; }
  if (preflight !== expected) return false;
  let count = 0, pass = 0, row = 0, column = 0, filter = 0, previousAbove = 0, adlerA = 1, adlerB = 0, adlerCount = 0;
  let priorRow = colorType === 3 ? new Uint8Array(passes[0]!.rowBytes - 1) : undefined;
  const inflater = new Inflate(chunk => {
    for (const value of chunk) {
      if (count >= expected || pass >= passes.length || (column === 0 && value > 4)) throw invalid;
      if (priorRow) {
        if (column === 0) { filter = value; previousAbove = 0; }
        else {
          const i = column - 1, above = priorRow[i]!, left = i ? priorRow[i - 1]! : 0, upperLeft = previousAbove;
          let predictor = 0;
          if (filter === 1) predictor = left;
          else if (filter === 2) predictor = above;
          else if (filter === 3) predictor = Math.floor((left + above) / 2);
          else if (filter === 4) {
            const p = left + above - upperLeft, a = Math.abs(p - left), b = Math.abs(p - above), c = Math.abs(p - upperLeft);
            predictor = a <= b && a <= c ? left : b <= c ? above : upperLeft;
          }
          const decoded = (value + predictor) & 255; priorRow[i] = decoded; previousAbove = above;
          const perByte = 8 / bitDepth, first = i * perByte, mask = (1 << bitDepth) - 1;
          for (let j = 0; j < perByte && first + j < passes[pass]!.width; j++) if (((decoded >>> (8 - bitDepth * (j + 1))) & mask) >= paletteEntries) throw invalid;
        }
      }
      adlerA += value; adlerB += adlerA;
      if (++adlerCount === 5552) { adlerA %= 65521; adlerB %= 65521; adlerCount = 0; }
      count++; if (++column === passes[pass]!.rowBytes) { column = 0; if (++row === passes[pass]!.rows) { row = 0; pass++; if (priorRow && pass < passes.length) priorRow = new Uint8Array(passes[pass]!.rowBytes - 1); } }
    }
  });
  try {
    for (let offset = 0; offset < raw.length; offset += 1024) { const stop = Math.min(raw.length, offset + 1024); inflater.push(raw.subarray(offset, stop), stop === raw.length); if ((offset / 1024) % 32 === 31) await checkpoint(); }
  } catch (error) { if (error instanceof PptxImportError || (error instanceof PackageInspectionError && error.code === "operation-aborted")) throw error; return false; }
  await checkpoint(); adlerA %= 65521; adlerB %= 65521;
  const checksum = new DataView(data.buffer, data.byteOffset + data.length - 4, 4).getUint32(0);
  return count === expected && pass === passes.length && (((adlerB << 16) | adlerA) >>> 0) === checksum;
}
