import test from "node:test";
import assert from "node:assert/strict";
import { inflateRawSync } from "node:zlib";

import { createZip } from "../dist/archive/create-zip.js";
import { corePropertiesXml } from "../dist/ooxml/package-parts.js";

test("OOXML helpers escape metadata independently of export orchestration", () => {
  assert.match(corePropertiesXml("Q1 & <ready>"), /Q1 &amp; &lt;ready&gt;/);
});

test("ZIP module emits a readable single-part package", () => {
  const bytes = Buffer.from(createZip([{ name: "hello.txt", data: new TextEncoder().encode("hello") }]));
  assert.equal(bytes.readUInt32LE(0), 0x04034b50);
  const nameLength = bytes.readUInt16LE(26);
  const compressedSize = bytes.readUInt32LE(18);
  assert.equal(bytes.readUInt16LE(10), 0);
  assert.equal(bytes.readUInt16LE(12), 0x0021);
  const compressed = bytes.subarray(30 + nameLength, 30 + nameLength + compressedSize);
  assert.equal(inflateRawSync(compressed).toString(), "hello");
  const centralOffset = 30 + nameLength + compressedSize;
  assert.equal(bytes.readUInt32LE(centralOffset), 0x02014b50);
  assert.equal(bytes.readUInt16LE(centralOffset + 12), 0);
  assert.equal(bytes.readUInt16LE(centralOffset + 14), 0x0021);
});
