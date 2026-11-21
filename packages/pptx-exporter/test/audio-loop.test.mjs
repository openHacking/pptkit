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
const source = readGeneratedZip(await readFile(join(root, "tests/fixtures/pptkit-mixed-av.pptx")));
const sound = source.get("ppt/media/sound.wav"), movie = source.get("ppt/media/movie.mp4");
const opts = { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" };
const attrs = (loop, rewind) => ({ ...(loop ? { repeatCount: "indefinite" } : {}), fill: rewind ? "remove" : "hold" });
function verify(path, expected) { execFileSync("python3", [join(root, "scripts/research/verify-audio-loop.py"), path, JSON.stringify(expected)]); }
function make(elements) {
  const p = createPresentation();
  p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "bytes", value: sound } });
  p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: { type: "bytes", value: movie } });
  p.addSlide({ elements: elements.map((x, i) => ({ id: `item-${i}`, box: { x: 50 * i, y: 0, width: 40, height: 40 }, ...x })) });
  return p;
}
const audio = (loop, rewindAfterPlayback, playbackStart) => ({ type: "audio", audioAssetId: "sound", ...(loop === undefined ? {} : { loop }), ...(rewindAfterPlayback === undefined ? {} : { rewindAfterPlayback }), ...(playbackStart === undefined ? {} : { playbackStart }) });

test("native and authored repeat and rewind are independent for each media instance", async () => {
  for (const [mode, loop, rewind] of [["loop", true, false], ["rewind", false, true], ["loop-rewind", true, true]]) verify(join(root, `tests/fixtures/powerpoint-audio-${mode}.pptx`), { 2: attrs(false, false), 3: attrs(false, false), 4: attrs(loop, rewind), 5: attrs(false, false) });
  const dir = await mkdtemp(join(tmpdir(), "pptkit-loop-"));
  try {
    for (const elements of [
      [audio(false, false), audio(true, false), audio(false, true), audio(true, true)],
      [audio(true, true, "object-click"), audio(false, true), audio(true, false, "object-click")],
      [{ type: "video", videoAssetId: "movie" }, audio(true, false), audio(false, true)],
      [{ type: "video", videoAssetId: "movie" }, audio(true, true, "object-click"), audio(false, false)],
    ]) {
      const p = make(elements), normalized = normalizePresentation(p).slides[0].elements;
      assert.deepEqual(normalized.filter(x => x.type === "audio").map(x => [x.loop, x.rewindAfterPlayback]), elements.filter(x => x.type === "audio").map(x => [x.loop, x.rewindAfterPlayback]));
      const path = join(dir, "out.pptx"), result = await generatePptx(p, opts); await writeFile(path, result.bytes);
      verify(path, Object.fromEntries(elements.flatMap((x, i) => x.type === "audio" ? [[String(i + 2), attrs(x.loop, x.rewindAfterPlayback)]] : [])));
      const entries = readGeneratedZip(result.bytes); assert.equal([...entries.keys()].filter(x => x.endsWith(".wav")).length, 1); assert.deepEqual(entries.get("ppt/media/sound.wav"), sound);
    }
    const out = join(dir, "sample.pptx"); execFileSync(process.execPath, [join(root, "scripts/research/generate-audio-loop.mjs"), out]);
    verify(out, { 2: attrs(false, false), 3: attrs(true, false), 4: attrs(false, true), 5: attrs(true, true) });
    assert.deepEqual(await readFile(out), await readFile(join(root, "tests/fixtures/pptkit-audio-loop.pptx")));
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test("false settings retain default bytes and invalid values fail before media I/O", async () => {
  assert.deepEqual((await generatePptx(make([audio()]), opts)).bytes, (await generatePptx(make([audio(false, false)]), opts)).bytes);
  for (const field of ["loop", "rewindAfterPlayback"]) for (const value of [null, 0, "true"]) {
    const p = createPresentation(); p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "url", value: "https://example.invalid/a.wav" } });
    p.addSlide({ elements: [{ ...audio(), [field]: value, box: { x: 0, y: 0, width: 40, height: 40 } }] });
    assert.ok(validatePresentation(p).some(x => x.code === (field === "loop" ? "invalid-audio-loop" : "invalid-audio-rewind")));
    let calls = 0; await assert.rejects(generatePptx(p, { audioMode: "wave-pcm-v1", resolveMedia() { calls++; } }), /Presentation validation failed/); assert.equal(calls, 0);
  }
});
