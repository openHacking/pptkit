import assert from "node:assert/strict";
import { constants, deflateRawSync, inflateRawSync } from "node:zlib";
import test from "node:test";
import { inspectZipArchive } from "../dist/index.js";
import { zip } from "./zip-fixture.mjs";

const inspect = (compressed, body = Buffer.alloc(0), options) => inspectZipArchive(zip([{ name: "part", body, compressed }]).bytes, options);
const invalid = (compressed, body) => assert.rejects(inspect(compressed, body), (error) => error.code === "archive-invalid");
function bits() {
  const values = [];
  return {
    value(value, width) { for (let i = 0; i < width; i++) values.push((value >>> i) & 1); return this; },
    code(value, width) { for (let i = width - 1; i >= 0; i--) values.push((value >>> i) & 1); return this; },
    fixed(symbol) {
      return symbol < 144 ? this.code(symbol + 0x30, 8) : symbol < 256 ? this.code(symbol - 144 + 0x190, 9) : symbol < 280 ? this.code(symbol - 256, 7) : this.code(symbol - 280 + 0xc0, 8);
    },
    bytes() { const bytes = Buffer.alloc(Math.ceil(values.length / 8)); values.forEach((value, i) => { bytes[i >>> 3] |= value << (i & 7); }); return bytes; },
  };
}
const order = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
function dynamic(literal, distance) {
  const writer = bits().value(1, 1).value(2, 2).value(literal.length - 257, 5).value(distance.length - 1, 5).value(15, 4);
  // Independent simple complete alphabet: symbols 0..15 use four-bit codes.
  for (const symbol of order) writer.value(symbol < 16 ? 4 : 0, 3);
  for (const length of [...literal, ...distance]) writer.code(length, 4);
  return writer;
}

test("final-block boundary rejects hidden bytes and concatenated streams but allows final-byte padding", async () => {
  const body = Buffer.from("boundary"), compressed = deflateRawSync(body);
  for (const tail of [Buffer.from([0]), Buffer.from([255]), deflateRawSync(Buffer.alloc(0)), Buffer.alloc(4096)]) await invalid(Buffer.concat([compressed, tail]), body);
  const padded = bits().value(1, 1).value(1, 2).fixed(256).bytes();
  padded[padded.length - 1] |= 0xfc;
  assert.equal(inflateRawSync(padded).length, 0);
  await inspect(padded);
});

test("truncated payloads and a missing final block reject despite plausible ZIP sizes and CRC", async () => {
  const body = Buffer.from("A sufficiently varied payload 1234567890".repeat(20));
  for (const strategy of [constants.Z_DEFAULT_STRATEGY, constants.Z_FIXED, constants.Z_HUFFMAN_ONLY]) {
    const compressed = deflateRawSync(body, { strategy });
    for (let cut = 0; cut < compressed.length; cut++) await invalid(compressed.subarray(0, cut), body);
  }
  await invalid(bits().value(0, 1).value(1, 2).fixed(256).bytes());
  await invalid(Buffer.from([1, 0, 0, 0, 0])); // Bad LEN/NLEN even though output is empty.
});

test("reserved symbols and back-references before history reject", async () => {
  await invalid(bits().value(1, 1).value(1, 2).fixed(286).fixed(256).bytes());
  await invalid(bits().value(1, 1).value(1, 2).fixed(257).code(0, 5).fixed(256).bytes());
  await invalid(bits().value(1, 1).value(1, 2).fixed(65).fixed(257).code(30, 5).fixed(256).bytes(), Buffer.from("A"));
  const valid = bits().value(1, 1).value(1, 2).fixed(65).fixed(285).code(0, 5).fixed(256).bytes();
  assert.deepEqual(inflateRawSync(valid), Buffer.alloc(259, 65));
  await inspect(valid, Buffer.alloc(259, 65));
});

test("dynamic trees validate empty-distance, single-symbol, missing-end, oversubscribed and incomplete alphabets", async () => {
  const empty = Array(257).fill(0); empty[256] = 1;
  const valid = dynamic(empty, [0]).code(0, 1).bytes();
  assert.equal(inflateRawSync(valid).length, 0); await inspect(valid);
  const missingEnd = [...empty]; missingEnd[256] = 0; missingEnd[65] = 1;
  await invalid(dynamic(missingEnd, [0]).code(0, 1).bytes());
  const over = [...empty]; over[65] = over[66] = 1;
  await invalid(dynamic(over, [0]).bytes());
  const incomplete = [...empty]; incomplete[256] = 2;
  await invalid(dynamic(incomplete, [0]).bytes());
  const needDistance = [...empty, 1];
  await invalid(dynamic(needDistance, [0]).code(1, 1).bytes());
  const codeOver = bits().value(1, 1).value(2, 2).value(0, 5).value(0, 5).value(0, 4);
  for (let i = 0; i < 4; i++) codeOver.value(1, 3);
  await invalid(codeOver.bytes());
});

test("dynamic code-length repeats reject missing predecessors and alphabet overruns", async () => {
  // Code alphabet has symbol 0 => 0 and repeat 16 => 1.
  const previous = bits().value(1, 1).value(2, 2).value(0, 5).value(0, 5).value(0, 4);
  for (const length of [1, 0, 0, 1]) previous.value(length, 3);
  await invalid(previous.code(1, 1).value(0, 2).bytes());
  // Complete code alphabet 0 => 0, repeat 18 => 1; two 138-zero runs exceed 258.
  const overrun = bits().value(1, 1).value(2, 2).value(0, 5).value(0, 5).value(0, 4);
  for (const length of [0, 0, 1, 1]) overrun.value(length, 3);
  await invalid(overrun.code(1, 1).value(127, 7).code(1, 1).value(127, 7).bytes());
});

test("seeded cross-decoder corpus covers stored/fixed/dynamic blocks, strategies and window sizes", async () => {
  let seed = 0x5eed1234;
  const random = () => { seed ^= seed << 13; seed ^= seed >>> 17; seed ^= seed << 5; return seed >>> 0; };
  const strategies = [constants.Z_DEFAULT_STRATEGY, constants.Z_FIXED, constants.Z_HUFFMAN_ONLY, constants.Z_RLE];
  const kinds = new Set();
  for (let index = 0; index < 80; index++) {
    const length = index < 4 ? index : 100 + random() % 70000;
    const body = Buffer.alloc(length);
    for (let i = 0; i < length; i++) body[i] = index % 3 === 0 ? random() & 255 : index % 3 === 1 ? i % 17 : random() % 5;
    const compressed = deflateRawSync(body, { level: index % 10, strategy: strategies[index % 4], windowBits: 9 + index % 7 });
    kinds.add((compressed[0] >>> 1) & 3);
    const decoded = inflateRawSync(compressed, { info: true });
    assert.equal(decoded.engine.bytesWritten, compressed.length); assert.deepEqual(decoded.buffer, body);
    assert.equal((await inspect(compressed, body)).expandedBytes, length);
  }
  assert.deepEqual([...kinds].sort(), [0, 1, 2]);
});

test("the structural scan cooperatively cancels even for streams of empty blocks", async () => {
  const writer = bits();
  for (let i = 0; i < 20000; i++) writer.value(i === 19999 ? 1 : 0, 1).value(1, 2).fixed(256);
  const controller = new AbortController();
  const result = inspect(writer.bytes(), Buffer.alloc(0), { signal: controller.signal });
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(result, (error) => error.code === "operation-aborted");
});

test("history crosses block boundaries and supports the full 32 KiB backward window", async () => {
  const history = Buffer.alloc(32768);
  for (let i = 0; i < history.length; i++) history[i] = i % 251;
  const header = Buffer.from([0, 0, 128, 255, 127]); // non-final stored block, LEN=32768
  const final = bits().value(1, 1).value(1, 2).fixed(257).code(29, 5).value(8191, 13).fixed(256).bytes();
  const compressed = Buffer.concat([header, history, final]), body = Buffer.concat([history, history.subarray(0, 3)]);
  assert.deepEqual(inflateRawSync(compressed), body);
  await inspect(compressed, body);
});

test("stalled-input budget applies to valid streams with no output progress", async () => {
  const writer = bits();
  for (let i = 0; i < 100; i++) writer.value(i === 99 ? 1 : 0, 1).value(1, 2).fixed(256);
  const compressed = writer.bytes();
  await inspect(compressed);
  await assert.rejects(inspect(compressed, Buffer.alloc(0), { limits: { maxStalledInputBytes: 16 } }), (error) => error.code === "archive-limit-exceeded");
});
