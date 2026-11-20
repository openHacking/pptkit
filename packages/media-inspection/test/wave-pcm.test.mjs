import test from "node:test";
import assert from "node:assert/strict";
import { inspectWavePcm } from "../dist/index.js";

function chunk(name, data) {
  const head = Buffer.alloc(8); head.write(name); head.writeUInt32LE(data.length, 4);
  return Buffer.concat([head, data, ...(data.length % 2 ? [Buffer.from([0])] : [])]);
}
function wave({ channels = 1, rate = 8000, bits = 8, frames = 80, extra = [] } = {}) {
  const fmt = Buffer.alloc(16), block = channels * bits / 8;
  fmt.writeUInt16LE(1, 0); fmt.writeUInt16LE(channels, 2); fmt.writeUInt32LE(rate, 4);
  fmt.writeUInt32LE(rate * block, 8); fmt.writeUInt16LE(block, 12); fmt.writeUInt16LE(bits, 14);
  const body = Buffer.concat([Buffer.from("WAVE"), ...extra, chunk("fmt ", fmt), chunk("data", Buffer.alloc(block * frames))]);
  const head = Buffer.alloc(8); head.write("RIFF"); head.writeUInt32LE(body.length, 4);
  return Buffer.concat([head, body]);
}

test("bounded WAVE inspection reports PCM frames without decoding", () => {
  for (const channels of [1, 2]) for (const bits of [8, 16]) {
    const source = wave({ channels, bits, extra: [chunk("JUNK", Buffer.from([3]))] });
    const input = Buffer.concat([Buffer.from([1, 2, 3]), source, Buffer.from([4])]).subarray(3, -1);
    const result = inspectWavePcm(input);
    assert.deepEqual(result, { format: "wave-pcm", channels, sampleRate: 8000, bitsPerSample: bits, frameCount: 80, durationSeconds: .01, dataOffset: 54, dataLength: channels * bits / 8 * 80 });
    assert.ok(Object.isFrozen(result));
  }
});

test("WAVE rejects damaged, ambiguous or unsupported structures", () => {
  const valid = wave(), cases = [];
  const mutate = (offset, value, size = 1) => { const out = Buffer.from(valid); if (size === 1) out[offset] = value; else if (size === 2) out.writeUInt16LE(value, offset); else out.writeUInt32LE(value, offset); return out; };
  cases.push(valid.subarray(0, 11), valid.subarray(0, -1), Buffer.concat([valid, Buffer.from([0])]), mutate(0, 0), mutate(20, 3, 2), mutate(22, 3, 2), mutate(24, 0, 4), mutate(28, 1, 4), mutate(32, 2, 2), mutate(34, 24, 2), mutate(40, 0xffffffff, 4));
  const duplicate = Buffer.concat([valid, chunk("data", Buffer.alloc(0))]); duplicate.writeUInt32LE(duplicate.length - 8, 4); cases.push(duplicate);
  const odd = wave({ channels: 2, bits: 16, frames: 1 }); odd.writeUInt32LE(3, 40); cases.push(odd);
  for (const bytes of cases) assert.throws(() => inspectWavePcm(bytes), TypeError);
  assert.throws(() => inspectWavePcm(valid, { maxBytes: valid.length - 1 }), RangeError);
  assert.throws(() => inspectWavePcm(valid, { maxChunks: 1 }), RangeError);
  assert.throws(() => inspectWavePcm(valid, { maxChunks: 4097 }), RangeError);
  const controller = new AbortController(); controller.abort(new Error("cancelled"));
  assert.throws(() => inspectWavePcm(valid, { signal: controller.signal }), /cancelled/);
});
