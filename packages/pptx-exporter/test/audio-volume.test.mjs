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
const fixture = readGeneratedZip(await readFile(join(root, "tests/fixtures/pptkit-mixed-av.pptx")));
const sounds = fixture.get("ppt/media/sound.wav"), movie = fixture.get("ppt/media/movie.mp4");
const opts = { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" };
const attrs = { low: { vol: "20000" }, medium: {}, high: { vol: "80000" }, mute: { mute: "1" } };
function verify(path, expected) {
  execFileSync("python3", [join(root, "scripts/research/verify-audio-volume.py"), path, JSON.stringify(expected)]);
}
function presentation(elements) {
  const p = createPresentation();
  p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "bytes", value: sounds } });
  p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: { type: "bytes", value: movie } });
  p.addSlide({ elements: elements.map((x, i) => ({ id: `item-${i}`, box: { x: i * 50, y: 0, width: 40, height: 40 }, ...x })) });
  return p;
}
const audio = (volume, start) => ({ type: "audio", audioAssetId: "sound", ...(volume === undefined ? {} : { playbackVolume: volume }), ...(start === undefined ? {} : { playbackStart: start }) });

test("native samples and generated independent instances have the expected volume nodes", async () => {
  for (const mode of ["low", "medium", "mute"]) verify(join(root, `tests/fixtures/powerpoint-audio-volume-${mode}.pptx`), { 2: attrs[mode] });
  const dir = await mkdtemp(join(tmpdir(), "pptkit-volume-"));
  try {
    for (const elements of [
      [audio("low"), audio("medium"), audio("high"), audio("mute")],
      [audio("mute", "object-click"), audio("low", "click-sequence"), audio("medium", "object-click"), audio("high")],
      [{ type: "video", videoAssetId: "movie" }, audio("low"), audio("medium"), audio("mute")],
      [{ type: "video", videoAssetId: "movie" }, audio("mute", "object-click"), audio("low")],
    ]) {
      const p = presentation(elements), normalized = normalizePresentation(p).slides[0].elements;
      assert.deepEqual(normalized.filter(x => x.type === "audio").map(x => x.playbackVolume), elements.filter(x => x.type === "audio").map(x => x.playbackVolume));
      const path = join(dir, "out.pptx"), result = await generatePptx(p, opts); await writeFile(path, result.bytes);
      const expected = Object.fromEntries(elements.flatMap((x, i) => x.type === "audio" ? [[String(i + 2), attrs[x.playbackVolume]]] : []));
      verify(path, expected);
      const entries = readGeneratedZip(result.bytes);
      assert.equal([...entries.keys()].filter(x => x.endsWith(".wav")).length, 1, "shared asset remains shared");
      assert.deepEqual(entries.get("ppt/media/sound.wav"), sounds);
    }
    const out = join(dir, "sample.pptx");
    execFileSync(process.execPath, [join(root, "scripts/research/generate-audio-volume.mjs"), out]);
    verify(out, { 2: attrs.low, 3: attrs.medium, 4: attrs.high, 5: attrs.mute });
    assert.deepEqual(await readFile(out), await readFile(join(root, "tests/fixtures/pptkit-audio-volume.pptx")));
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test("default and explicit high retain identical bytes; bad values fail before I/O", async () => {
  assert.deepEqual((await generatePptx(presentation([audio()]), opts)).bytes, (await generatePptx(presentation([audio("high")]), opts)).bytes);
  for (const bad of ["silent", 0, null, false]) {
    const p = createPresentation();
    p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "url", value: "https://example.invalid/a.wav" } });
    p.addSlide({ elements: [{ ...audio(bad), box: { x: 0, y: 0, width: 40, height: 40 } }] });
    assert.ok(validatePresentation(p).some(x => x.code === "invalid-audio-playback-volume"));
    let calls = 0;
    await assert.rejects(generatePptx(p, { audioMode: "wave-pcm-v1", resolveMedia() { calls++; } }), /Presentation validation failed/);
    assert.equal(calls, 0);
  }
});
