import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { inspectIsoBmff, inspectMp4MovieDuration } from "../dist/index.js";

function box(type, payload = Buffer.alloc(0), extended = false) {
  const header = Buffer.alloc(extended ? 16 : 8);
  header.writeUInt32BE(extended ? 1 : header.length + payload.length);
  header.write(type, 4, 4, "latin1");
  if (extended) header.writeBigUInt64BE(BigInt(header.length + payload.length), 8);
  return Buffer.concat([header, payload]);
}
const ftyp = () => box("ftyp", Buffer.concat([Buffer.from("isom"), Buffer.alloc(4), Buffer.from("isommp42")]));

test("reports bounded top-level framing without interpreting media", async () => {
  const bytes = await readFile(new URL("../../../tests/fixtures/h264-mp4-powerpoint.pptx", import.meta.url));
  const mp4 = readGeneratedZip(bytes).get("ppt/media/media1.mp4");
  const view = Buffer.concat([Buffer.from([1, 2]), mp4, Buffer.from([3])]).subarray(2, -1);
  const result = inspectIsoBmff(view);
  assert.equal(result.majorBrand, "isom");
  assert.deepEqual(result.boxes.map(x => [x.type, x.offset, x.size]), [["ftyp", 0, 32], ["moov", 32, 1163], ["free", 1195, 8], ["mdat", 1203, 42325]]);
  assert.ok(Object.isFrozen(result) && Object.isFrozen(result.boxes) && Object.isFrozen(result.boxes[0]) && Object.isFrozen(result.compatibleBrands));
  const extended = inspectIsoBmff(Buffer.concat([ftyp(), box("free", Buffer.from([7]), true), Buffer.from([0, 0, 0, 0, 109, 100, 97, 116, 4])]));
  assert.deepEqual(extended.compatibleBrands, ["isom", "mp42"]);
  assert.deepEqual(extended.boxes.map(x => [x.type, x.headerSize, x.size]), [["ftyp", 8, 24], ["free", 16, 17], ["mdat", 8, 9]]);
});

test("rejects truncated, ambiguous and over-budget framing", () => {
  const valid = Buffer.concat([ftyp(), box("mdat", Buffer.from([1]))]);
  const cases = [Buffer.alloc(0), valid.subarray(0, 7), valid.subarray(0, -1), box("free"), Buffer.concat([ftyp(), ftyp()]), box("ftyp", Buffer.alloc(7)), box("ftyp", Buffer.alloc(9)), Buffer.from([0, 0, 0, 1, 102, 114, 101, 101]), Buffer.from([0, 0, 0, 7, 102, 114, 101, 101])];
  const oversized = box("free", Buffer.alloc(0), true); oversized.writeUInt32BE(1, 8); cases.push(Buffer.concat([ftyp(), oversized]));
  for (const bytes of cases) assert.throws(() => inspectIsoBmff(bytes), TypeError);
  assert.throws(() => inspectIsoBmff(valid, { maxBytes: valid.length - 1 }), RangeError);
  assert.throws(() => inspectIsoBmff(valid, { maxBoxes: 1 }), RangeError);
  assert.throws(() => inspectIsoBmff(valid, { maxBoxes: 4097 }), RangeError);
  assert.throws(() => inspectIsoBmff(box("ftyp", Buffer.alloc(8 + 4097 * 4))), RangeError);
  const controller = new AbortController(); controller.abort(new Error("cancelled"));
  assert.throws(() => inspectIsoBmff(valid, { signal: controller.signal }), /cancelled/);
});

test("captures caller bytes before visiting boxes", () => {
  const bytes = Buffer.concat([ftyp(), box("mdat", Buffer.from([1]))]);
  const expected = inspectIsoBmff(bytes); let checks = 0;
  const signal = { get aborted() { if (++checks === 2) bytes[4] = 0; return false; } };
  assert.deepEqual(inspectIsoBmff(bytes, { signal }), expected);
  assert.equal(bytes[4], 0);
});

function movieHeader(scale = 1000, duration = 1250) {
  const data = Buffer.alloc(100); data.writeUInt32BE(scale, 12); data.writeUInt32BE(duration, 16); return box("mvhd", data);
}
function movie(...children) { return Buffer.concat([ftyp(), box("moov", Buffer.concat(children)), box("mdat", Buffer.from([1]))]); }

test("reads the bounded version-zero movie clock without interpreting tracks", async () => {
  assert.equal(inspectMp4MovieDuration(movie(movieHeader(3000, 3703))), 1234);
  const mp4 = readGeneratedZip(await readFile(new URL("../../../tests/fixtures/h264-mp4-powerpoint.pptx", import.meta.url))).get("ppt/media/media1.mp4");
  assert.equal(inspectMp4MovieDuration(mp4), 2000);
  const zeroSized = movieHeader(); zeroSized.writeUInt32BE(0, 0);
  assert.equal(inspectMp4MovieDuration(movie(zeroSized)), 1250);
  const extended = box("mvhd", movieHeader().subarray(8), true);
  assert.equal(inspectMp4MovieDuration(movie(extended)), 1250);
});

test("rejects ambiguous, incomplete, fragmented and over-budget movie clocks", () => {
  const header = movieHeader(), version = Buffer.from(header); version[8] = 1;
  const flags = Buffer.from(header); flags[11] = 1;
  const truncated = Buffer.from(header); truncated.writeUInt32BE(500, 0);
  const shortExtended = Buffer.from([0, 0, 0, 1, 102, 114, 101, 101]);
  const cases = [movie(), movie(header, header), movie(version), movie(flags), movie(movieHeader(0, 1)), movie(movieHeader(1, 0)), movie(movieHeader(1, 0xffffffff)), movie(movieHeader(1, 0xfffffffe)), movie(box("mvhd", Buffer.alloc(99))), movie(truncated), movie(shortExtended), Buffer.concat([movie(header), box("moof")]), Buffer.concat([ftyp(), box("moov", header), box("mdat")]), Buffer.concat([movie(header), box("moov", header)])];
  for (const value of cases) assert.throws(() => inspectMp4MovieDuration(value), TypeError);
  const valid = movie(header);
  assert.throws(() => inspectMp4MovieDuration(valid, { maxBytes: valid.length - 1 }), RangeError);
  assert.throws(() => inspectMp4MovieDuration(movie(...Array.from({ length: 4097 }, () => box("free")), header)), RangeError);
  const controller = new AbortController(); controller.abort(new Error("cancelled"));
  assert.throws(() => inspectMp4MovieDuration(valid, { signal: controller.signal }), /cancelled/);
});

test("movie clock uses a captured input snapshot", () => {
  const input = movie(movieHeader()); let checks = 0;
  const signal = { get aborted() { if (++checks === 2) input.fill(0); return false; }, throwIfAborted() {} };
  assert.equal(inspectMp4MovieDuration(input, { signal }), 1250);
  assert.equal(input[4], 0);
});
