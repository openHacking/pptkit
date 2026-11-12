import test from "node:test";
import assert from "node:assert/strict";
import { inflateRawSync } from "node:zlib";

import { createZip } from "../dist/archive/create-zip.js";
import { corePropertiesXml } from "../dist/ooxml/package-parts.js";
import { relationshipsXml } from "../dist/ooxml/relationships.js";

test("relationship attributes escape query strings and XML delimiters", () => {
  const xml = relationshipsXml([{ id: "rId1", type: "https://example.com/type", target: 'https://example.com/?a=1&b="<value>"&c=\'quoted\'', targetMode: "External" }]);
  assert.ok(xml.includes('Target="https://example.com/?a=1&amp;b=&quot;&lt;value&gt;&quot;&amp;c=&apos;quoted&apos;"'));
  assert.match(xml, /TargetMode="External"/);
});

test("OOXML helpers escape metadata independently of export orchestration", () => {
  assert.match(corePropertiesXml({
    title: "Q1 & <ready>",
    author: "PPTKit",
    language: "en-US",
    keywords: [],
    revision: 1,
  }), /Q1 &amp; &lt;ready&gt;/);
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
