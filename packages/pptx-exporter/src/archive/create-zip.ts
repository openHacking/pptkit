import { deflateRawSync } from "node:zlib";
import type { ZipPart } from "../types/internal.js";

// ZIP stores dates in the MS-DOS format. Zero is not a valid date because
// month and day are one-based, so use the earliest representable date.
const DOS_DATE_1980_01_01 = 0x0021;

export function createZip(parts: ZipPart[]): Buffer {
  const local: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const part of parts) {
    const compressed = deflateRawSync(part.data);
    const crc = crc32(part.data);
    const name = Buffer.from(part.name, "utf8");
    const header = Buffer.alloc(30 + name.length);
    header.writeUInt32LE(0x04034b50, 0); header.writeUInt16LE(20, 4); header.writeUInt16LE(0, 6); header.writeUInt16LE(8, 8); header.writeUInt16LE(0, 10); header.writeUInt16LE(DOS_DATE_1980_01_01, 12); header.writeUInt32LE(crc, 14); header.writeUInt32LE(compressed.length, 18); header.writeUInt32LE(part.data.length, 22); header.writeUInt16LE(name.length, 26); header.writeUInt16LE(0, 28); name.copy(header, 30);
    local.push(header, compressed);
    const entry = Buffer.alloc(46 + name.length);
    entry.writeUInt32LE(0x02014b50, 0); entry.writeUInt16LE(20, 4); entry.writeUInt16LE(20, 6); entry.writeUInt16LE(0, 8); entry.writeUInt16LE(8, 10); entry.writeUInt16LE(0, 12); entry.writeUInt16LE(DOS_DATE_1980_01_01, 14); entry.writeUInt32LE(crc, 16); entry.writeUInt32LE(compressed.length, 20); entry.writeUInt32LE(part.data.length, 24); entry.writeUInt16LE(name.length, 28); entry.writeUInt16LE(0, 30); entry.writeUInt16LE(0, 32); entry.writeUInt16LE(0, 34); entry.writeUInt16LE(0, 36); entry.writeUInt32LE(0, 38); entry.writeUInt32LE(offset, 42); name.copy(entry, 46);
    central.push(entry); offset += header.length + compressed.length;
  }
  const centralData = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(central.length, 8); end.writeUInt16LE(central.length, 10); end.writeUInt32LE(centralData.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...local, centralData, end]);
}

function crc32(data: Buffer): number {
  let crc = 0xffffffff;
  for (const byte of data) { crc ^= byte; for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0); }
  return (crc ^ 0xffffffff) >>> 0;
}
