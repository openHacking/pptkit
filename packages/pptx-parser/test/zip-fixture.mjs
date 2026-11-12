import { deflateRawSync } from "node:zlib";
export function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
export function zip(entries, options = {}) {
  const locals = [], centrals = [], offsets = [];
  let position = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name), body = Buffer.from(entry.body ?? ""), method = entry.method ?? 8;
    const compressed = entry.compressed ?? (method === 0 ? body : deflateRawSync(body));
    const flags = entry.flags ?? (0x800 | (entry.descriptor ? 8 : 0));
    const crc = entry.crc ?? crc32(body), expanded = entry.expanded ?? body.length;
    const local = Buffer.alloc(30 + name.length); local.writeUInt32LE(0x04034b50); local.writeUInt16LE(20, 4); local.writeUInt16LE(flags, 6); local.writeUInt16LE(method, 8);
    if (!entry.descriptor) { local.writeUInt32LE(crc, 14); local.writeUInt32LE(compressed.length, 18); local.writeUInt32LE(expanded, 22); }
    local.writeUInt16LE(name.length, 26); name.copy(local, 30);
    let descriptor = Buffer.alloc(0);
    if (entry.descriptor) {
      const signed = entry.descriptor === "signed"; descriptor = Buffer.alloc(signed ? 16 : 12);
      if (signed) descriptor.writeUInt32LE(0x08074b50);
      const base = signed ? 4 : 0; descriptor.writeUInt32LE(crc, base); descriptor.writeUInt32LE(compressed.length, base + 4); descriptor.writeUInt32LE(expanded, base + 8);
    }
    const central = Buffer.alloc(46 + name.length); central.writeUInt32LE(0x02014b50); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(flags, 8); central.writeUInt16LE(method, 10); central.writeUInt32LE(crc, 16); central.writeUInt32LE(compressed.length, 20); central.writeUInt32LE(expanded, 24); central.writeUInt16LE(name.length, 28); central.writeUInt32LE(position, 42); name.copy(central, 46);
    offsets.push({ local: position, data: position + local.length, descriptor: position + local.length + compressed.length });
    locals.push(local, compressed, descriptor); centrals.push(central); position += local.length + compressed.length + descriptor.length;
  }
  const directory = position;
  const ordered = options.reverseCentral ? [...centrals].reverse() : centrals;
  const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10); end.writeUInt32LE(centrals.reduce((sum, bytes) => sum + bytes.length, 0), 12); end.writeUInt32LE(directory, 16);
  const comment = Buffer.from(options.comment ?? ""); end.writeUInt16LE(comment.length, 20);
  return { bytes: Buffer.concat([...locals, ...ordered, end, comment]), directory, offsets };
}
