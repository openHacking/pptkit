import test from "node:test";
import assert from "node:assert/strict";
import { readFile, writeFile, mkdtemp, rm } from "node:fs/promises";
import { execFileSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { createPresentation, normalizePresentation, validatePresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const root = fileURLToPath(new URL("../../../", import.meta.url));
const movie = readGeneratedZip(await readFile(join(root, "tests/fixtures/pptkit-mixed-av.pptx"))).get("ppt/media/movie.mp4");
const video = (start, end, more = {}) => ({ type: "video", videoAssetId: "movie", ...(start === undefined ? {} : { trimStartMs: start }), ...(end === undefined ? {} : { trimEndMs: end }), ...more });
function make(elements, external = false) { const p = createPresentation(); p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: external ? { type: "url", value: "https://example.invalid/video.mp4" } : { type: "bytes", value: movie } }); p.addSlide({ elements: elements.map((e, i) => ({ id: `video-${i}`, box: { x: 50 * i, y: 0, width: 40, height: 40 }, ...e })) }); return p; }
const expected = (start = 0, end = 0) => ({ trim: start || end ? { st: String(start), end: String(end) } : null, duration: String(2000 - start - end) });
const verify = (path, values) => execFileSync("python3", [join(root, "scripts/research/verify-video-trim.py"), path, JSON.stringify(values)]);
test("video trims retain source bytes, per-object timing and automatic sequence offsets", async () => {
  verify(join(root, "tests/fixtures/powerpoint-video-trim.pptx"), { 2: expected(150, 50), 3: expected() });
  const native = readGeneratedZip(await readFile(join(root, "tests/fixtures/powerpoint-video-trim.pptx")));
  assert.deepEqual(native.get("ppt/media/media1.mp4"), movie);
  const dir = await mkdtemp(join(tmpdir(), "pptkit-video-trim-"));
  try {
    const elements = [video(150, 50, { playbackStart: "automatic" }), video(0, 100, { playbackStart: "automatic" })];
    const p = make(elements); assert.deepEqual(normalizePresentation(p).slides[0].elements.map(x => [x.trimStartMs, x.trimEndMs]), [[150, 50], [0, 100]]);
    const result = await generatePptx(p, { videoMode: "iso-bmff-v1" }), path = join(dir, "out.pptx");
    await writeFile(path, result.bytes); verify(path, { 2: expected(150, 50), 3: expected(0, 100) });
    const archive = readGeneratedZip(result.bytes); assert.deepEqual(archive.get("ppt/media/movie.mp4"), movie);
    assert.equal([...archive.keys()].filter(x => x.endsWith(".mp4")).length, 1);
    const groups = [{ begin: [[null, "indefinite", null], ["onBegin", "0", "2"]], effects: [{ target: "2", type: "afterEffect", delay: "0", duration: "1800" }, { target: "3", type: "afterEffect", delay: "1800", duration: "1900" }] }];
    execFileSync("python3", [join(root, "scripts/research/verify-audio-auto.py"), path, JSON.stringify(groups)]);
    execFileSync(process.execPath, [join(root, "scripts/research/generate-video-trim.mjs"), path]);
    verify(path, { 2: expected(150, 50), 3: expected() });
    assert.deepEqual(await readFile(path), await readFile(join(root, "tests/fixtures/pptkit-video-trim.pptx")));
    assert.deepEqual((await generatePptx(make([video()]), { videoMode: "iso-bmff-v1" })).bytes, (await generatePptx(make([video(0, 0)]), { videoMode: "iso-bmff-v1" })).bytes);
  } finally { await rm(dir, { recursive: true, force: true }); }
});
test("video trim validates each shared use and rejects before external resolution where possible", async () => {
  for (const key of ["trimStartMs", "trimEndMs"]) for (const value of [-1, 0.5, NaN, Infinity, null, "1", 2147483648]) {
    const p = make([video(undefined, undefined, { [key]: value })], true); assert.ok(validatePresentation(p).some(x => x.code === "invalid-video-trim"));
    let calls = 0; await assert.rejects(generatePptx(p, { videoMode: "iso-bmff-v1", resolveMedia() { calls++; } }), /validation failed/i); assert.equal(calls, 0);
  }
  for (const elements of [[video(2000, 0)], [video(1000, 1000)], [video(0, 0), video(1999, 1)]]) {
    await assert.rejects(generatePptx(make(elements), { videoMode: "iso-bmff-v1" }), /trim must leave at least one millisecond/);
    let calls = 0; await assert.rejects(generatePptx(make(elements, true), { videoMode: "iso-bmff-v1", resolveMedia() { calls++; return { bytes: movie }; } }), /trim must leave at least one millisecond/); assert.equal(calls, 1);
  }
});
