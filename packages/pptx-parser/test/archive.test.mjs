import assert from "node:assert/strict";
import test from "node:test";
import { importPptx, inspectZipArchive, PackageInspectionError, DEFAULT_ZIP_LIMITS } from "../dist/index.js";
import { crc32, zip } from "./zip-fixture.mjs";
import { structureFixture, encodeStructure } from "./structure-fixture.mjs";
const rejects = (bytes, code, options) => assert.rejects(inspectZipArchive(bytes, options), (error) => error instanceof PackageInspectionError && error.code === code);

test("stored/deflate records, directories, Unicode names and comments are verified", async () => {
  const { bytes } = zip([{ name: "ppt/", method: 0 }, { name: "ppt/slides/slide1.xml", body: "123456789", method: 0 }, { name: "ppt/media/中文.txt", body: "hello" }], { comment: "Comment", reverseCentral: true });
  const before = Buffer.from(bytes), result = await inspectZipArchive(bytes);
  assert.equal(crc32(Buffer.from("123456789")), 0xcbf43926);
  assert.equal(result.entries.find((entry) => entry.name.endsWith("slide1.xml")).crc32, 0xcbf43926);
  assert.equal(result.expandedBytes, 14); assert.equal(result.inputBytes, bytes.length);
  assert.deepEqual(bytes, before); assert.equal(result.entries.length, 3);
  assert.ok(Object.isFrozen(DEFAULT_ZIP_LIMITS));
});

test("both signed and unsigned ZIP32 data descriptors match central metadata", async () => {
  for (const descriptor of ["signed", "unsigned"]) {
    const fixture = zip([{ name: "part", body: "streamed data", descriptor }]);
    assert.equal((await inspectZipArchive(fixture.bytes)).expandedBytes, 13);
    fixture.bytes[fixture.offsets[0].descriptor + (descriptor === "signed" ? 4 : 0)] ^= 1;
    await rejects(fixture.bytes, "archive-invalid");
  }
});

test("input is detached before the first asynchronous yield", async () => {
  const { bytes } = zip([{ name: "part", body: "original" }]);
  const result = inspectZipArchive(bytes); bytes.fill(0);
  assert.equal((await result).expandedBytes, 8);
});

test("CRC corruption, local filename mismatch and local size mismatches fail", async () => {
  const fixture = zip([{ name: "part", body: "hello", method: 0 }]);
  let bytes = Buffer.from(fixture.bytes); bytes[fixture.offsets[0].data] ^= 1; await rejects(bytes, "archive-crc-mismatch");
  bytes = Buffer.from(fixture.bytes); bytes[30] ^= 1; await rejects(bytes, "archive-invalid");
  bytes = Buffer.from(fixture.bytes); bytes.writeUInt32LE(99, 22); await rejects(bytes, "archive-invalid");
});

test("unsafe paths, percent aliases, case aliases and canonical Unicode aliases reject", async () => {
  for (const name of ["../a", "/a", "C:/a", "a\\b", "a//b", "a/./b", "a/%2e%2e/b", "a%00b", "a%252fb", "a%", "a#b", ""]) await rejects(zip([{ name, method: 0 }]).bytes, "archive-path-invalid");
  for (const names of [["a", "a"], ["a", "%61"], ["A", "a"], ["é", "e\u0301"], ["a", "a/"]]) await rejects(zip(names.map((name) => ({ name, method: 0 }))).bytes, "archive-duplicate-path");
});

test("unsupported ZIP features and invalid UTF-8 fail explicitly", async () => {
  for (const flags of [1, 0x40, 0x1000]) await rejects(zip([{ name: "a", flags, method: 0 }]).bytes, "archive-unsupported");
  await rejects(zip([{ name: "a", method: 99, compressed: Buffer.from("x") }]).bytes, "archive-unsupported");
  await rejects(zip([{ name: "中文", flags: 0, method: 0 }]).bytes, "archive-unsupported");
  const fixture = zip([{ name: "aa", method: 0 }]); fixture.bytes[30] = fixture.bytes[fixture.directory + 46] = 0xff;
  await rejects(fixture.bytes, "archive-path-invalid");
  const large = zip([{ name: "a", method: 0 }]); large.bytes.writeUInt32LE(0xffffffff, large.directory + 24); await rejects(large.bytes, "archive-unsupported");
});

test("overlap, hidden local bytes, trailing junk, truncation and inconsistent directories reject", async () => {
  const fixture = zip([{ name: "a", body: "a" }, { name: "b", body: "b" }]);
  let bytes = Buffer.from(fixture.bytes); bytes.writeUInt32LE(0, fixture.directory + 47 + 42); await rejects(bytes, "archive-invalid");
  await rejects(Buffer.concat([fixture.bytes, Buffer.from("junk")]), "archive-invalid");
  for (let cut = 0; cut < fixture.bytes.length; cut++) await assert.rejects(inspectZipArchive(fixture.bytes.subarray(0, cut)), PackageInspectionError);
  bytes = Buffer.from(fixture.bytes); bytes.writeUInt16LE(1, bytes.length - 22 + 8); await rejects(bytes, "archive-unsupported");
});

test("declared input, entry, expanded-total and ratio budgets fail before decoding", async () => {
  const { bytes } = zip([{ name: "a", body: "a".repeat(1000) }, { name: "b", body: "b".repeat(1000) }]);
  for (const limits of [{ maxInputBytes: bytes.length - 1 }, { maxEntries: 1 }, { maxEntryBytes: 999 }, { maxExpandedBytes: 1999 }, { maxCompressionRatio: 1 }]) await rejects(bytes, "archive-limit-exceeded", { limits });
});

test("a forged expanded size cannot bypass streaming output budgets", async () => {
  const fixture = zip([{ name: "bomb", body: Buffer.alloc(2 * 1024 * 1024), expanded: 10 }]);
  await rejects(fixture.bytes, "archive-limit-exceeded", { limits: { maxEntryBytes: 100 } });
  await rejects(fixture.bytes, "archive-invalid", { limits: { maxEntryBytes: 4 * 1024 * 1024 } });
});

test("invalid deflate streams and empty compressed payloads reject", async () => {
  await rejects(zip([{ name: "part", compressed: Buffer.from([7]), body: "" }]).bytes, "archive-invalid");
  await rejects(zip([{ name: "part", compressed: Buffer.alloc(0), body: "" }]).bytes, "archive-invalid");
});

test("abort is observed before work and after cooperative chunk yields", async () => {
  const early = new AbortController(); early.abort(); await rejects(new Uint8Array(), "operation-aborted", { signal: early.signal });
  const { bytes } = zip([{ name: "large", method: 0, body: Buffer.alloc(256 * 1024) }]);
  const controller = new AbortController(); const pending = inspectZipArchive(bytes, { signal: controller.signal });
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(pending, (error) => error.code === "operation-aborted");
});

test("malformed budgets reject as configuration errors and empty ZIP remains only an archive", async () => {
  for (const limits of [{ maxEntries: 0 }, { maxInputBytes: Infinity }, { maxExpandedBytes: -1 }, { maxEntryBytes: 1.5 }]) await assert.rejects(inspectZipArchive(zip([]).bytes, { limits }), RangeError);
  assert.deepEqual((await inspectZipArchive(zip([]).bytes)).entries, []);
});


test("file/directory hierarchy conflicts and special filesystem entries reject", async () => {
  await rejects(zip([{ name: "a", method: 0 }, { name: "a/b", method: 0 }]).bytes, "archive-path-invalid");
  const fixture = zip([{ name: "link", body: "target", method: 0 }]);
  fixture.bytes.writeUInt16LE(0x314, fixture.directory + 4);
  fixture.bytes.writeUInt32LE((0xa000 << 16) >>> 0, fixture.directory + 38);
  await rejects(fixture.bytes, "archive-unsupported");
});

test("directory names agree with Unix file types without interpreting foreign host mode bits", async () => {
  const archive = (name, host, mode, dosDirectory = false) => {
    const fixture = zip([{ name, method: 0 }]);
    fixture.bytes.writeUInt16LE((host << 8) | 20, fixture.directory + 4);
    fixture.bytes.writeUInt32LE(((mode << 16) | (dosDirectory ? 16 : 0)) >>> 0, fixture.directory + 38);
    return fixture.bytes;
  };
  for (const host of [3, 19]) {
    for (const bytes of [archive('folder/', host, 0x8000), archive('file', host, 0x4000), archive('file', host, 0x8000, true)]) await rejects(bytes, 'archive-path-invalid');
    for (const type of [0xa000, 0xc000, 0x1000, 0x2000, 0x6000]) await rejects(archive('entry', host, type), 'archive-unsupported');
  }
  for (const [bytes, directory] of [[archive('file', 3, 0x8000), false], [archive('folder/', 3, 0x4000, true), true], [archive('folder/', 3, 0), true], [archive('file', 19, 0x8000), false], [archive('folder/', 19, 0x4000), true], [archive('folder/', 19, 0), true], [archive('file', 0, 0x4000), false], [archive('file', 0, 0xa000), false]]) {
    const result = await inspectZipArchive(bytes); assert.equal(result.entries[0].directory, directory);
  }
  await rejects(archive('folder/', 3, 0xa000), 'archive-unsupported');
});

test("OS X entry types gate full PPTX import and valid no-op saves retain archive bytes", async () => {
  const original = Buffer.from(encodeStructure(structureFixture()));
  const changeType = mode => {
    const bytes = Buffer.from(original), end = bytes.length - 22;
    assert.equal(bytes.readUInt32LE(end), 0x06054b50);
    const central = bytes.readUInt32LE(end + 16);
    assert.equal(bytes.readUInt32LE(central), 0x02014b50);
    bytes.writeUInt16LE((19 << 8) | 20, central + 4);
    bytes.writeUInt32LE((mode << 16) >>> 0, central + 38);
    return bytes;
  };
  const valid = changeType(0x8000), link = changeType(0xa000), wrongDirectory = changeType(0x4000);
  for (const projection of ['structure', 'content']) {
    const session = await importPptx(valid, { projection });
    try {
      const prepared = await session.prepareSave({ mode: 'no-op' }); assert.equal(prepared.status, 'ready');
      assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), valid);
    } finally { session.dispose(); }
    await assert.rejects(importPptx(link, { projection }), { code: 'import-unsupported' });
    await assert.rejects(importPptx(wrongDirectory, { projection }), { code: 'import-invalid-package' });
  }
});

test("caller options cannot replace a captured abort signal during asynchronous work", async () => {
  const controller = new AbortController();
  const options = { signal: controller.signal };
  const pending = inspectZipArchive(zip([{ name: "large", method: 0, body: Buffer.alloc(128 * 1024) }]).bytes, options);
  options.signal = new AbortController().signal; controller.abort();
  await assert.rejects(pending, (error) => error.code === "operation-aborted");
});

test("deep paths and adjacent sibling names do not hide file/directory conflicts", async () => {
  const deep = "a/".repeat(16000) + "leaf";
  assert.equal((await inspectZipArchive(zip([{ name: deep, method: 0 }]).bytes)).entries[0].name, deep);
  for (const names of [["a", "a-", "a/b"], ["a/b/c", "a/b-", "a/b"], ["a/b/", "a-", "a"]]) await rejects(zip(names.map((name) => ({ name, method: 0 }))).bytes, "archive-path-invalid");
  const names = ["a", "a-", "a-b/c", "ab/c", "ab/"];
  assert.equal((await inspectZipArchive(zip(names.map((name) => ({ name, method: 0 }))).bytes)).entries.length, names.length);
});
