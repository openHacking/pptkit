import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { importPptx } from "../dist/index.js";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../../pptx-exporter/dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const { zipSync, unzipSync, strFromU8, strToU8 } = createRequire(import.meta.url)("fflate");
const options = { projection: "content", audioProjection: "wave-pcm-v1" };
const sample = name => readFile(new URL(`../../../tests/fixtures/${name}.pptx`, import.meta.url));
const starts = s => s.document.slides[0].elements.filter(x => x.type === "audio").map(x => x.playbackStart);

async function makeAudio(modes) {
  const parts = readGeneratedZip(await sample("pptkit-mixed-av")), p = createPresentation();
  p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "bytes", value: parts.get("ppt/media/sound.wav") } });
  p.addSlide({ elements: modes.map((mode, i) => ({ type: "audio", audioAssetId: "sound", playbackStart: mode, ...(i === 0 ? { trimStartMs: 100, trimEndMs: 50 } : {}), box: { x: i * 50, y: 0, width: 40, height: 40 } })) });
  return (await generatePptx(p, { audioMode: "wave-pcm-v1" })).bytes;
}

test("native and generated complete main audio sequences project in order", async () => {
  for (const [name, expected] of [["powerpoint-single-audio-auto", ["automatic"]], ["powerpoint-two-audio-auto", ["automatic", "automatic"]], ["pptkit-two-audio-auto", ["automatic", "automatic"]], ["pptkit-audio-trim", ["click-sequence", "object-click", "click-sequence", "object-click"]]]) {
    const bytes = await sample(name), s = await importPptx(bytes, options);
    try { assert.deepEqual(starts(s), expected); const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes)); }
    finally { s.dispose(); }
  }
  for (const modes of [["click-sequence", "automatic", "automatic"], ["automatic", "object-click", "automatic", "click-sequence"]]) {
    let bytes = await makeAudio(modes);
    for (let i = 0; i < 3; i++) { const s = await importPptx(bytes, options); try { assert.deepEqual(starts(s), modes); bytes = (await generatePptx(s.document, { audioMode: "wave-pcm-v1" })).bytes; } finally { s.dispose(); } }
  }
});

function modify(bytes, change) { const entries = unzipSync(bytes), key = "ppt/slides/slide1.xml", before = strFromU8(entries[key]), after = change(before); assert.notEqual(before, after); entries[key] = strToU8(after); return zipSync(entries); }
test("uncertain main timing leaves the entire main sequence unmapped", async () => {
  const original = await makeAudio(["automatic", "automatic"]);
  const changes = [
    s => s.replace('evt="onBegin" delay="0"', 'evt="onClick" delay="0"'),
    s => s.replace('restart="never" nodeType="tmRoot"', 'restart="always" nodeType="tmRoot"'),
    s => s.replace('delay="850"', 'delay="849"'),
    s => s.replace('dur="850" fill="hold"', 'dur="849" fill="hold"'),
    s => s.replace('nodeType="afterEffect"', 'nodeType="withEffect"'),
    s => s.replace('presetClass="mediacall"', 'presetClass="entr"'),
    s => s.replace(/(<p:pic>.*?<\/p:pic>)(<p:pic>.*?<\/p:pic>)/, '$2$1'),
    s => s.replace('id="6" dur="850"', 'id="02" dur="850"'),
    s => s.replace('<p:cMediaNode vol="80000"><p:cTn id="10"', '<p:cMediaNode vol="80000"><p:cTn repeatCount="indefinite" id="10"'),
    s => s.replace('<p:cMediaNode vol="80000"><p:cTn id="10"', '<p:cMediaNode vol="80000"><p:cTn repeatCount="2000" id="10"'),
    s => s.replace('spid="2"/></p:tgtEl></p:cBhvr></p:cmd>', 'spid="3"/></p:tgtEl></p:cBhvr></p:cmd>'),
  ];
  for (const change of changes) { const bytes = modify(original, change), s = await importPptx(bytes, options); try { assert.deepEqual(starts(s), [undefined, undefined]); const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, bytes); } finally { s.dispose(); } }
});

test("unprojected main video prevents partial automatic timing projection", async () => {
  const bytes = await sample("powerpoint-audio-start-auto"), s = await importPptx(bytes, options);
  try { assert.deepEqual(starts(s), [undefined, undefined]); assert.ok(s.report.diagnostics.some(x => x.code === "media-playback-unmapped")); } finally { s.dispose(); }
});


const mixedOptions = { ...options, videoProjection: "iso-bmff-v1" };
async function makeMixed() {
  const parts = readGeneratedZip(await sample("pptkit-mixed-av")), p = createPresentation();
  p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "bytes", value: parts.get("ppt/media/sound.wav") } });
  p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: { type: "bytes", value: parts.get("ppt/media/movie.mp4") } });
  p.addSlide({ elements: [
    { type: "video", videoAssetId: "movie", box: { x: 0, y: 0, width: 40, height: 40 } },
    { type: "audio", audioAssetId: "sound", playbackStart: "automatic", box: { x: 50, y: 0, width: 40, height: 40 } },
    { type: "audio", audioAssetId: "sound", playbackStart: "object-click", box: { x: 100, y: 0, width: 40, height: 40 } },
    { type: "video", videoAssetId: "movie", box: { x: 150, y: 0, width: 40, height: 40 } },
    { type: "audio", audioAssetId: "sound", playbackStart: "automatic", box: { x: 200, y: 0, width: 40, height: 40 } },
  ] });
  return (await generatePptx(p, { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" })).bytes;
}

test("complete mixed media sequences project audio starts across video effects", async () => {
  let bytes = await sample("powerpoint-audio-start-auto");
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, mixedOptions);
    try {
      assert.deepEqual(starts(s), ["click-sequence", "automatic"]);
      const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
      bytes = (await generatePptx(s.document, { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" })).bytes;
    } finally { s.dispose(); }
  }
  bytes = await makeMixed();
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, mixedOptions);
    try { assert.deepEqual(starts(s), ["automatic", "object-click", "automatic"]); bytes = (await generatePptx(s.document, { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" })).bytes; }
    finally { s.dispose(); }
  }
});

test("ambiguous video timing blocks the whole mixed main sequence", async () => {
  const original = await sample("powerpoint-audio-start-auto");
  const changes = [
    s => s.replace('cmd="togglePause"', 'cmd="playFrom(0.0)"'),
    s => s.replace('presetID="2" presetClass="mediacall"', 'presetID="3" presetClass="mediacall"'),
    s => s.replace('dur="2000" fill="hold"', 'dur="1999" fill="hold"'),
    s => s.replace('<p:video><p:cMediaNode vol="80000">', '<p:video><p:cMediaNode vol="50000">'),
    s => s.replace('<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId1"/>', '<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId1"><p14:trim st="1"/></p14:media>'),
    s => s.replace('evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="2"', 'evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="3"'),
  ];
  for (const change of changes) {
    const bytes = modify(original, change), s = await importPptx(bytes, mixedOptions);
    try { assert.deepEqual(starts(s), [undefined, undefined]); const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, bytes); }
    finally { s.dispose(); }
  }
  const independent = modify(await makeMixed(), xml => xml.replace('cmd="togglePause"', 'cmd="stop"'));
  const mixed = await importPptx(independent, mixedOptions);
  try { assert.deepEqual(starts(mixed), [undefined, "object-click", undefined]); } finally { mixed.dispose(); }
  const parts = unzipSync(original), video = parts['ppt/media/media1.mp4'];
  const index = Buffer.from(video).indexOf(Buffer.from('mvhd')); assert.ok(index > 0); video[index + 4] = 1;
  const s = await importPptx(zipSync(parts), mixedOptions);
  try { assert.deepEqual(starts(s), [undefined, undefined]); } finally { s.dispose(); }
});


test("native reordered automatic video remains unmapped instead of changing playback order", async () => {
  const bytes = await sample("powerpoint-video-start-auto"), session = await importPptx(bytes, mixedOptions);
  try {
    assert.deepEqual(session.document.slides[0].elements.map(x => x.type), ["video", "video", "audio", "audio"]);
    assert.deepEqual(starts(session), [undefined, undefined]);
    assert.ok(session.report.diagnostics.some(x => x.code === "media-playback-unmapped"));
    const plan = await session.prepareSave({ mode: "no-op" });
    assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
  } finally { session.dispose(); }
});
