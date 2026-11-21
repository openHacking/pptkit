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
const parts = readGeneratedZip(await readFile(join(root, "tests/fixtures/pptkit-mixed-av.pptx")));
const sound = parts.get("ppt/media/sound.wav"), movie = parts.get("ppt/media/movie.mp4");
const options = { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" };
const auto = (target, delay = 0, duration = 1000) => ({ target: String(target), type: "afterEffect", delay: String(delay), duration: String(duration) });
const click = (target, duration = 1000) => ({ target: String(target), type: "clickEffect", delay: "0", duration: String(duration) });
const group = (effects, begin = false) => ({ begin: [[null, "indefinite", null], ...(begin ? [["onBegin", "0", "2"]] : [])], effects });
const verify = (path, expected) => execFileSync("python3", [join(root, "scripts/research/verify-audio-auto.py"), path, JSON.stringify(expected)]);
const audio = (playbackStart, extra = {}) => ({ type: "audio", audioAssetId: "sound", playbackStart, ...extra });
function make(elements) {
  const p = createPresentation(); p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "bytes", value: sound } });
  p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: { type: "bytes", value: movie } });
  p.addSlide({ elements: elements.map((x, i) => ({ id: `item-${i}`, box: { x: i * 50, y: 0, width: 40, height: 40 }, ...x })) }); return p;
}

test("automatic media uses native main sequence grouping and bounded delays", async () => {
  verify(join(root, "tests/fixtures/powerpoint-single-audio-auto.pptx"), [group([auto(2)], true)]);
  verify(join(root, "tests/fixtures/powerpoint-two-audio-auto.pptx"), [group([auto(2), auto(3, 1000)], true)]);
  const dir = await mkdtemp(join(tmpdir(), "pptkit-auto-"));
  try {
    const cases = [
      [[audio("automatic")], [group([auto(2)], true)]],
      [[audio("automatic", { trimStartMs: 100, trimEndMs: 50 }), audio("automatic")], [group([auto(2, 0, 850), auto(3, 850)], true)]],
      [[audio("click-sequence"), audio("automatic"), audio("automatic")], [group([click(2), auto(3, 1000), auto(4, 2000)])]],
      [[audio("automatic"), audio("click-sequence"), audio("automatic")], [group([auto(2)], true), group([click(3), auto(4, 1000)])]],
      [[{ type: "video", videoAssetId: "movie" }, audio("object-click"), audio("automatic")], [group([click(2, 2000), auto(4, 2000)])]],
    ];
    for (const [elements, expected] of cases) {
      const p = make(elements), output = join(dir, "out.pptx"); assert.equal(normalizePresentation(p).slides[0].elements.filter(x => x.type === "audio").at(-1).playbackStart, "automatic");
      await writeFile(output, (await generatePptx(p, options)).bytes); verify(output, expected);
    }
    const out = join(dir, "sample.pptx");
    for (const [arg, name, expected] of [["--single", "pptkit-single-audio-auto.pptx", [group([auto(2)], true)]], [undefined, "pptkit-two-audio-auto.pptx", [group([auto(2), auto(3, 1000)], true)]]]) {
      execFileSync(process.execPath, [join(root, "scripts/research/generate-audio-auto.mjs"), out, ...(arg ? [arg] : [])]); verify(out, expected); assert.deepEqual(await readFile(out), await readFile(join(root, "tests/fixtures", name)));
    }
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test("automatic after an indefinitely looping main effect is rejected", async () => {
  await assert.rejects(generatePptx(make([audio("click-sequence", { loop: true }), audio("automatic")]), options), /cannot follow looping audio/);
  await assert.rejects(generatePptx(make([audio("automatic", { loop: true }), audio("object-click"), audio("automatic")]), options), /cannot follow looping audio/);
  await generatePptx(make([audio("automatic", { loop: true }), audio("click-sequence")]), options);
});


test("automatic video uses ordered native media calls with independent pauses", async () => {
  const video = playbackStart => ({ type: "video", videoAssetId: "movie", playbackStart });
  const dir = await mkdtemp(join(tmpdir(), "pptkit-video-auto-"));
  try {
    const cases = [
      [[video("automatic")], [group([auto(2, 0, 2000)], true)]],
      [[video("click-sequence"), video("automatic")], [group([click(2, 2000), auto(3, 2000, 2000)])]],
      [[audio("click-sequence"), video("automatic"), audio("automatic")], [group([click(2), auto(3, 1000, 2000), auto(4, 3000)])]],
      [[video("automatic"), audio("object-click"), video("automatic")], [group([auto(2, 0, 2000), auto(4, 2000, 2000)], true)]],
    ];
    for (const [elements, expected] of cases) {
      const p = make(elements);
      assert.equal(normalizePresentation(p).slides[0].elements.find(x => x.type === "video" && x.playbackStart === "automatic").playbackStart, "automatic");
      const path = join(dir, "out.pptx"); await writeFile(path, (await generatePptx(p, options)).bytes); verify(path, expected);
    }
    const generated = join(dir, "generated.pptx");
    execFileSync(process.execPath, [join(root, "scripts/research/generate-video-auto.mjs"), generated]);
    assert.deepEqual(await readFile(generated), await readFile(join(root, "tests/fixtures/pptkit-video-auto.pptx")));
    verify(generated, [group([auto(2, 0, 2000), auto(3, 2000, 2000)], true)]);
    const base = make([video(undefined)]), clickVideo = make([video("click-sequence")]);
    assert.deepEqual((await generatePptx(base, options)).bytes, (await generatePptx(clickVideo, options)).bytes);
    const invalid = make([video("with-previous")]);
    assert.ok(validatePresentation(invalid).some(x => x.code === "invalid-video-playback-start"));
    await assert.rejects(generatePptx(invalid, options), /validation failed/i);
    await assert.rejects(generatePptx(make([audio("click-sequence", { loop: true }), video("automatic")]), options), /cannot follow looping audio/);
  } finally { await rm(dir, { recursive: true, force: true }); }
});
