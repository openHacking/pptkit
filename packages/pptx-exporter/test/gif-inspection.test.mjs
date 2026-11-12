import test from "node:test";
import assert from "node:assert/strict";
import { inspectGif } from "../dist/index.js";
import { inspectGif as inspectNode } from "../dist/node.js";

const bytes = (...parts) => Uint8Array.from(parts.flat(Infinity));
const header = [...Buffer.from("GIF89a"), 2, 0, 2, 0, 0x80, 0, 0, 0, 0, 0, 255, 255, 255];
const loop = [0x21, 0xff, 11, ...Buffer.from("NETSCAPE2.0"), 3, 1, 0, 0, 0];
const control = (delay, flags = 0) => [0x21, 0xf9, 4, flags, delay & 255, delay >> 8, 1, 0];
const frame = (flags = 0) => [0x2c, 0, 0, 0, 0, 2, 0, 2, 0, flags, 2, 2, 0x44, 1, 0];
const sample = () => bytes(header, loop, control(0, 9), frame(), [0x21, 0xfe, 2, 1, 2, 0], control(7, 8), frame(0x40), 0x3b);

test("graphic-control user input follows its next graphic across comments and is consumed once", () => {
  const comment = [0x21, 0xfe, 1, 65, 0];
  const text = [0x21, 1, 12, ...Array(12).fill(0), 1, 65, 0];
  const input = bytes(header, control(7, 2), comment, frame(), frame(), control(0, 2), text, frame(), control(0, 28), frame(), 0x3b);
  const result = inspectGif(input);
  assert.deepEqual(result.frames.map(f => [f.userInputFlag, f.delayCentiseconds, f.disposal]), [[true, 7, 0], [undefined, 0, 0], [undefined, 0, 0], [undefined, 0, 7]]);
  assert.deepEqual(inspectNode(input), result);
  assert.equal(inspectGif(bytes(header, control(0, 0xe0), frame(), 0x3b)).frames[0].reservedControlBits, 0xe0);
});

test("reports frames, delays, transparency, disposal, interlace and loop without decoding", () => {
  const input = sample(), original = input.slice(), result = inspectGif(input);
  assert.deepEqual(result, { version: "89a", width: 2, height: 2, loopCount: 0, frames: [
    { left: 0, top: 0, width: 2, height: 2, delayCentiseconds: 0, disposal: 2, transparentIndex: 1, interlaced: false },
    { left: 0, top: 0, width: 2, height: 2, delayCentiseconds: 7, disposal: 2, interlaced: true },
  ] });
  assert.deepEqual(inspectNode(input), result); assert.deepEqual(input, original);
  const offset = new Uint8Array(input.length + 4); offset.set(input, 2);
  assert.deepEqual(inspectGif(offset.subarray(2, -2)), result);
  result.frames[0].width = 999; assert.equal(inspectGif(input).frames[0].width, 2);
});

test("inspects an independently encoded two-frame GIF", () => {
  // Pillow 12.1.1 encoded and decoded red/blue 2x2 frames with delays 0/70 ms and repeat count 2.
  const value = "R0lGODlhAgACAIEAAP8AAAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwECAAAh+QQIAAAAACwAAAAAAgACAAAIBgABCAQQEAAh+QQIBwAAACwAAAAAAgACAIEAAP8AAAAAAAAAAAAIBgABCAQQEAA7";
  const result = inspectGif(Uint8Array.from(Buffer.from(value, "base64")));
  assert.equal(result.loopCount, 2);
  assert.deepEqual(result.frames.map(f => f.delayCentiseconds), [0, 7]);
  assert.deepEqual(result.frames.map(f => [f.width, f.height, f.disposal]), [[2, 2, 2], [2, 2, 2]]);
});

test("handles static GIF87a, local color table and control scope across plain text", () => {
  const local = [0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0x80, 0, 0, 0, 255, 255, 255, 2, 1, 0x44, 0];
  const text = [0x21, 1, 12, ...Array(12).fill(0), 1, 65, 0];
  const input = bytes([...Buffer.from("GIF87a"), 1, 0, 1, 0, 0, 0, 0], control(5), text, local, 0x3b);
  assert.deepEqual(inspectGif(input), { version: "87a", width: 1, height: 1, plainTextBlocks: 1, frames: [
    { left: 0, top: 0, width: 1, height: 1, delayCentiseconds: 0, disposal: 0, interlaced: false },
  ] });
});

test("rejects truncation, malformed blocks, duplicate controls and invalid geometry", () => {
  const input = sample();
  for (let length = 0; length < input.length; length++) assert.throws(() => inspectGif(input.subarray(0, length)), undefined, `prefix ${length}`);
  for (const bad of [bytes(input, 0), bytes(header, control(1), control(2), frame(), 0x3b), bytes(header, 0x77, 0x3b),
    bytes(header, [0x2c, 1, 0, 0, 0, 2, 0, 2, 0, 0], 2, 0, 0x3b), bytes(header, loop, loop, frame(), 0x3b)]) assert.throws(() => inspectGif(bad), /Invalid GIF/);
});

test("enforces finite input, frame and block budgets, option types, and cancellation", () => {
  const input = sample();
  assert.deepEqual(inspectGif(input, { maxBytes: input.length, maxFrames: 2 }), inspectGif(input));
  assert.throws(() => inspectGif(input, { maxBytes: input.length - 1 }), /maxBytes/);
  assert.throws(() => inspectGif(input, { maxFrames: 1 }), /maxFrames/);
  assert.throws(() => inspectGif(input, { maxBlocks: 2 }), /maxBlocks/);
  for (const key of ["maxBytes", "maxFrames", "maxBlocks"]) for (const value of [0, -1, 1.5, Infinity, NaN, null]) assert.throws(() => inspectGif(input, { [key]: value }), RangeError);
  assert.throws(() => inspectGif([]), TypeError);
  const controller = new AbortController(), reason = new Error("cancelled"); controller.abort(reason);
  assert.throws(() => inspectGif(input, { signal: controller.signal }), error => error === reason);
});

test("reports skipped visual and unknown extension blocks separately from comments", () => {
  const text = [0x21, 1, 12, ...Array(12).fill(0), 1, 65, 0];
  const app = [0x21, 0xff, 11, ...Buffer.from("UNKNOWN1234"), 1, 5, 0];
  const unknown = [0x21, 0xee, 1, 42, 0], comment = [0x21, 0xfe, 1, 65, 0];
  const result = inspectGif(bytes(header, text, app, unknown, comment, frame(), 0x3b));
  assert.equal(result.plainTextBlocks, 1); assert.equal(result.unknownExtensionBlocks, 2);
  assert.equal(result.frames.length, 1);
  const badLoop = [0x21, 0xff, 11, ...Buffer.from("NETSCAPE2.0"), 3, 1, 0, 0, 1, 9, 0];
  assert.throws(() => inspectGif(bytes(header, badLoop, frame(), 0x3b)), /unexpected loop extension data/);
});
