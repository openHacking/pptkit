import assert from "node:assert/strict";
import { inflateRawSync } from "node:zlib";

// Test-only reader for the ZIP32 archives emitted by PPTKit, not an input parser.
// The central directory is checked independently of local headers so truncated or
// inconsistent output cannot pass simply because an expected XML string exists.
export function readGeneratedZip(input) {
  const bytes = Buffer.from(input);
  assert.ok(bytes.length >= 22, "ZIP end record is present");
  const end = bytes.length - 22;
  assert.equal(bytes.readUInt32LE(end), 0x06054b50, "ZIP ends with an uncommented EOCD");
  const count = bytes.readUInt16LE(end + 10);
  const directory = bytes.readUInt32LE(end + 16);
  assert.equal(directory + bytes.readUInt32LE(end + 12), end, "central directory covers the archive tail");
  const entries = new Map();
  let cursor = directory;
  for (let index = 0; index < count; index++) {
    assert.equal(bytes.readUInt32LE(cursor), 0x02014b50, "central entry signature");
    const method = bytes.readUInt16LE(cursor + 10);
    const compressed = bytes.readUInt32LE(cursor + 20);
    const uncompressed = bytes.readUInt32LE(cursor + 24);
    const nameLength = bytes.readUInt16LE(cursor + 28);
    const extraLength = bytes.readUInt16LE(cursor + 30);
    const commentLength = bytes.readUInt16LE(cursor + 32);
    const name = bytes.subarray(cursor + 46, cursor + 46 + nameLength).toString("utf8");
    assert.ok(!entries.has(name), `unique part ${name}`);
    const local = bytes.readUInt32LE(cursor + 42);
    assert.equal(bytes.readUInt32LE(local), 0x04034b50);
    assert.equal(bytes.readUInt16LE(local + 8), method);
    assert.equal(bytes.readUInt32LE(local + 18), compressed);
    assert.equal(bytes.readUInt32LE(local + 22), uncompressed);
    const localNameLength = bytes.readUInt16LE(local + 26);
    assert.equal(bytes.subarray(local + 30, local + 30 + localNameLength).toString("utf8"), name);
    const start = local + 30 + localNameLength + bytes.readUInt16LE(local + 28);
    assert.ok(start + compressed <= directory, `part ${name} is inside local data`);
    assert.ok(method === 0 || method === 8, `supported ZIP method ${method}`);
    const data = bytes.subarray(start, start + compressed);
    const decoded = method === 8 ? inflateRawSync(data) : Buffer.from(data);
    assert.equal(decoded.length, uncompressed);
    entries.set(name, decoded);
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  assert.equal(cursor, end, "all central entries consumed");
  return entries;
}
