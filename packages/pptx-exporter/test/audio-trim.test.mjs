import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { execFileSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { createPresentation, normalizePresentation, validatePresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const root = fileURLToPath(new URL("../../../", import.meta.url));
const parts = readGeneratedZip(await readFile(join(root, "tests/fixtures/pptkit-mixed-av.pptx")));
const sound = parts.get("ppt/media/sound.wav"), movie = parts.get("ppt/media/movie.mp4");
const options = { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" };
const expected = (start = 0, end = 0) => ({ trim: start || end ? { st: String(start), end: String(end) } : null, duration: String(1000 - start - end) });
function verify(path, values) { execFileSync("python3", [join(root, "scripts/research/verify-audio-trim.py"), path, JSON.stringify(values)]); }
function make(elements, external = false) {
  const p = createPresentation(); p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: external ? { type: "url", value: "https://example.invalid/sound.wav" } : { type: "bytes", value: sound } });
  p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: { type: "bytes", value: movie } });
  p.addSlide({ elements: elements.map((x, i) => ({ id: `item-${i}`, box: { x: i * 50, y: 0, width: 40, height: 40 }, ...x })) }); return p;
}
const audio = (start, end, more = {}) => ({ type: "audio", audioAssetId: "sound", ...(start === undefined ? {} : { trimStartMs: start }), ...(end === undefined ? {} : { trimEndMs: end }), ...more });

test("native and authored trims use per-object offsets and duration", async () => {
  verify(join(root, "tests/fixtures/powerpoint-audio-trim.pptx"), { 2: expected(), 3: expected(), 4: expected(100, 50), 5: expected() });
  const dir = await mkdtemp(join(tmpdir(), "pptkit-trim-"));
  try {
    for (const elements of [
      [audio(0, 0), audio(100, 0), audio(100, 50), audio(0, 200)],
      [audio(100, 50, { playbackStart: "object-click", loop: true }), audio(0, 0), audio(999, 0)],
      [{ type: "video", videoAssetId: "movie" }, audio(100, 50), audio(0, 200)],
      [{ type: "video", videoAssetId: "movie" }, audio(100, 50, { playbackStart: "object-click" }), audio(0, 0)],
    ]) {
      const p = make(elements), normalized = normalizePresentation(p).slides[0].elements;
      assert.deepEqual(normalized.filter(x => x.type === "audio").map(x => [x.trimStartMs, x.trimEndMs]), elements.filter(x => x.type === "audio").map(x => [x.trimStartMs, x.trimEndMs]));
      const out = join(dir, "out.pptx"), result = await generatePptx(p, options); await writeFile(out, result.bytes);
      verify(out, Object.fromEntries(elements.flatMap((x, i) => x.type === "audio" ? [[String(i + 2), expected(x.trimStartMs, x.trimEndMs)]] : [])));
      const entries = readGeneratedZip(result.bytes); assert.equal([...entries.keys()].filter(x => x.endsWith(".wav")).length, 1); assert.deepEqual(entries.get("ppt/media/sound.wav"), sound);
    }
    const out = join(dir, "sample.pptx"); execFileSync(process.execPath, [join(root, "scripts/research/generate-audio-trim.mjs"), out]);
    verify(out, { 2: expected(), 3: expected(100, 0), 4: expected(100, 50), 5: expected(0, 200) });
    assert.deepEqual(await readFile(out), await readFile(join(root, "tests/fixtures/pptkit-audio-trim.pptx")));
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test("trim validation rejects invalid values and each shared instance is bounded", async () => {
  assert.deepEqual((await generatePptx(make([audio()]), options)).bytes, (await generatePptx(make([audio(0, 0)]), options)).bytes);
  for (const key of ["trimStartMs", "trimEndMs"]) for (const value of [-1, 0.5, NaN, Infinity, null, "1", 2147483648]) {
    const p = make([audio(undefined, undefined, { [key]: value })], true); assert.ok(validatePresentation(p).some(x => x.code === "invalid-audio-trim"));
    let calls = 0; await assert.rejects(generatePptx(p, { ...options, resolveMedia() { calls++; } }), /Presentation validation failed/); assert.equal(calls, 0);
  }
  for (const elements of [[audio(1000, 0)], [audio(500, 500)], [audio(0, 0), audio(999, 1)]]) {
    await assert.rejects(generatePptx(make(elements), options), /trim must leave at least one millisecond/);
    let calls = 0; await assert.rejects(generatePptx(make(elements, true), { ...options, resolveMedia() { calls++; return { bytes: sound }; } }), /trim must leave at least one millisecond/); assert.equal(calls, 1);
  }
});
