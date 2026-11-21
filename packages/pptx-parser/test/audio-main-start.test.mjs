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
    s => s.replace('<p:video><p:cMediaNode vol="80000">', '<p:video><p:cMediaNode vol="60000">'),
    s => s.replace('<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId1"/>', '<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId1"><p14:trim st="1"/></p14:media>'),
    s => s.replace('evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="2"', 'evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="3"'),
  ];
  for (const change of changes) {
    const bytes = modify(original, change), s = await importPptx(bytes, mixedOptions);
    try { assert.deepEqual(starts(s), [undefined, undefined]); assert.deepEqual(s.document.slides[0].elements.filter(x => x.type === "video").map(x => x.playbackStart), [undefined, undefined]); const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, bytes); }
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
    assert.deepEqual(session.document.slides[0].elements.filter(x => x.type === "video").map(x => x.playbackStart), [undefined, undefined]);
    assert.ok(session.report.diagnostics.some(x => x.code === "media-playback-unmapped"));
    const plan = await session.prepareSave({ mode: "no-op" });
    assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
  } finally { session.dispose(); }
});


test("ordered automatic video starts survive video-only and mixed fresh exports", async () => {
  const parts = readGeneratedZip(await sample("pptkit-mixed-av"));
  for (const mixed of [false, true]) {
    const p = createPresentation();
    p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: { type: "bytes", value: parts.get("ppt/media/movie.mp4") } });
    p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "bytes", value: parts.get("ppt/media/sound.wav") } });
    p.addSlide({ elements: [
      { type: "video", videoAssetId: "movie", playbackStart: "automatic", box: { x: 0, y: 0, width: 40, height: 40 } },
      ...(mixed ? [{ type: "audio", audioAssetId: "sound", playbackStart: "automatic", box: { x: 50, y: 0, width: 40, height: 40 } }] : []),
      { type: "video", videoAssetId: "movie", playbackStart: "click-sequence", box: { x: 100, y: 0, width: 40, height: 40 } },
      { type: "video", videoAssetId: "movie", playbackStart: "automatic", box: { x: 150, y: 0, width: 40, height: 40 } },
    ] });
    let bytes = (await generatePptx(p, { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" })).bytes;
    for (let round = 0; round < 3; round++) {
      const s = await importPptx(bytes, mixed ? mixedOptions : { projection: "content", videoProjection: "iso-bmff-v1" });
      try {
        assert.deepEqual(s.document.slides[0].elements.filter(x => x.type === "video").map(x => x.playbackStart), ["automatic", "click-sequence", "automatic"]);
        if (mixed) assert.deepEqual(starts(s), ["automatic"]);
        const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
        bytes = (await generatePptx(s.document, { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" })).bytes;
      } finally { s.dispose(); }
    }
  }
});


test("native and generated independent video clicks preserve mixed start modes", async () => {
  const native = await sample("powerpoint-video-object-click"), session = await importPptx(native, mixedOptions);
  try {
    assert.deepEqual(session.document.slides[0].elements.map(x => x.playbackStart), ["object-click", "click-sequence", "click-sequence", "click-sequence"]);
    const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(native));
  } finally { session.dispose(); }
  const parts = readGeneratedZip(await sample("pptkit-mixed-av"));
  for (const mixed of [false, true]) {
    const p = createPresentation();
    p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: { type: "bytes", value: parts.get("ppt/media/movie.mp4") } });
    p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "bytes", value: parts.get("ppt/media/sound.wav") } });
    p.addSlide({ elements: [
      { type: "video", videoAssetId: "movie", playbackStart: "object-click", box: { x: 0, y: 0, width: 40, height: 40 } },
      ...(mixed ? [{ type: "audio", audioAssetId: "sound", playbackStart: "automatic", box: { x: 50, y: 0, width: 40, height: 40 } }, { type: "video", videoAssetId: "movie", playbackStart: "click-sequence", box: { x: 100, y: 0, width: 40, height: 40 } }] : []),
      { type: "video", videoAssetId: "movie", playbackStart: "object-click", box: { x: 150, y: 0, width: 40, height: 40 } },
    ] });
    let bytes = (await generatePptx(p, { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" })).bytes;
    for (let round = 0; round < 3; round++) {
      const s = await importPptx(bytes, mixed ? mixedOptions : { projection: "content", videoProjection: "iso-bmff-v1" });
      try {
        assert.deepEqual(s.document.slides[0].elements.filter(x => x.type === "video").map(x => x.playbackStart), mixed ? ["object-click", "click-sequence", "object-click"] : ["object-click", "object-click"]);
        if (mixed) assert.deepEqual(starts(s), ["automatic"]);
        bytes = (await generatePptx(s.document, { audioMode: "wave-pcm-v1", videoMode: "iso-bmff-v1" })).bytes;
      } finally { s.dispose(); }
    }
  }
});

test("independent video start is not guessed from damaged pause or media nodes", async () => {
  const native = await sample("powerpoint-video-object-click");
  const changes = [
    s => s.replaceAll('cmd="togglePause"', 'cmd="stop"'),
    s => s.replace('evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="2"', 'evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="3"'),
    s => s.replace('<p:video><p:cMediaNode vol="80000"><p:cTn id="28"', '<p:video><p:cMediaNode vol="60000"><p:cTn id="28"'),
    s => s.replace('<p:cTn id="28" fill="hold" display="0"', '<p:cTn id="28" repeatCount="2000" fill="hold" display="0"'),
    s => s.replace('restart="never" nodeType="tmRoot"', 'restart="always" nodeType="tmRoot"'),
    s => s.replace(/(<p:video><p:cMediaNode vol="80000"><p:cTn id="28".*?<\/p:video>)/, '$1$1'),
    s => s.replaceAll('cmd="togglePause"', 'cmd="togglePause" extra="1"'),
  ];
  for (const change of changes) {
    const bytes = modify(native, change), s = await importPptx(bytes, mixedOptions);
    try { assert.equal(s.document.slides[0].elements.find(x => x.type === "video").playbackStart, undefined); } finally { s.dispose(); }
  }
});


test("native and generated video trim projects shared bytes and source-relative offsets", async () => {
  for (const name of ["powerpoint-video-trim", "pptkit-video-trim"]) {
    let bytes = await sample(name);
    for (let round = 0; round < 3; round++) {
      const s = await importPptx(bytes, { projection: "content", videoProjection: "iso-bmff-v1" });
      try {
        const videos = s.document.slides[0].elements.filter(x => x.type === "video");
        assert.equal(videos.length, 2); assert.equal(videos[0].videoAssetId, videos[1].videoAssetId);
        assert.deepEqual(videos.map(x => [x.trimStartMs, x.trimEndMs, x.playbackStart]), [[150, 50, "automatic"], [undefined, undefined, "automatic"]]);
        const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
        bytes = (await generatePptx(s.document, { videoMode: "iso-bmff-v1" })).bytes;
      } finally { s.dispose(); }
    }
  }
});

test("invalid video trim does not infer timing or trust external sentinel", async () => {
  const generated = await sample("pptkit-video-trim");
  const changes = [
    s => s.replace('st="150" end="50"', 'st="150.5" end="50"'),
    s => s.replace('st="150" end="50"', 'st="1950" end="50"'),
    s => s.replace('st="150" end="50"', 'st="150" end="50" extra="1"'),
    s => s.replace('delay="1800"', 'delay="1801"'),
    s => s.replace('dur="1800"', 'dur="1801"'),
  ];
  for (const change of changes) {
    const bytes = modify(generated, change), s = await importPptx(bytes, { projection: "content", videoProjection: "iso-bmff-v1" });
    try { assert.deepEqual(s.document.slides[0].elements.filter(x => x.type === "video").map(x => x.playbackStart), [undefined, undefined]); const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, bytes); } finally { s.dispose(); }
  }
  const native = await sample("powerpoint-video-trim"), damaged = modify(native, s => s.replace('st="150" end="50"', 'st="150.5" end="50"'));
  const session = await importPptx(damaged, { projection: "content", videoProjection: "iso-bmff-v1" });
  try { assert.equal(session.document.slides[0].elements.filter(x => x.type === "video").length, 1); assert.ok(session.report.videoSources.some(x => x.externalTarget === "NULL")); } finally { session.dispose(); }
});


test("video object-click remains independent with a valid trim", async () => {
  const parts = readGeneratedZip(await sample("pptkit-mixed-av")), p = createPresentation();
  p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: { type: "bytes", value: parts.get("ppt/media/movie.mp4") } });
  p.addSlide({ elements: [{ type: "video", videoAssetId: "movie", playbackStart: "object-click", trimStartMs: 150, trimEndMs: 50, box: { x: 0, y: 0, width: 40, height: 40 } }] });
  let bytes = (await generatePptx(p, { videoMode: "iso-bmff-v1" })).bytes;
  for (let i = 0; i < 3; i++) { const s = await importPptx(bytes, { projection: "content", videoProjection: "iso-bmff-v1" }); try { const [v] = s.document.slides[0].elements; assert.deepEqual([v.playbackStart, v.trimStartMs, v.trimEndMs], ["object-click", 150, 50]); bytes = (await generatePptx(s.document, { videoMode: "iso-bmff-v1" })).bytes; } finally { s.dispose(); } }
});

test("video settings preserve native mute and volume independently and survive fresh exports", async () => {
  for (const [name, muted] of [["powerpoint-video-settings-low", false], ["powerpoint-video-settings-mute", true]]) {
    const bytes = await sample(name), s = await importPptx(bytes, { projection: "content", videoProjection: "iso-bmff-v1" });
    try {
      const videos = s.document.slides[0].elements.filter(x => x.type === "video");
      assert.deepEqual(videos.map(x => [x.playbackVolume, x.muted, x.loop, x.rewindAfterPlayback, x.playbackStart]), [["low", muted, true, true, undefined], ["high", false, false, false, undefined]]);
      const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { s.dispose(); }
  }
  const parts = readGeneratedZip(await sample("pptkit-mixed-av")), p = createPresentation();
  p.registerAsset({ id: "movie", kind: "video", mimeType: "video/mp4", source: { type: "bytes", value: parts.get("ppt/media/movie.mp4") } });
  const expected = [["low", true, true, true, "object-click"], ["medium", false, false, false, "click-sequence"], ["high", false, false, true, "automatic"]];
  p.addSlide({ elements: expected.map(([playbackVolume, muted, loop, rewindAfterPlayback, playbackStart], i) => ({ type: "video", videoAssetId: "movie", playbackVolume, muted, loop, rewindAfterPlayback, playbackStart, box: { x: 50 * i, y: 0, width: 40, height: 40 } })) });
  let bytes = (await generatePptx(p, { videoMode: "iso-bmff-v1" })).bytes;
  for (let i = 0; i < 3; i++) {
    const s = await importPptx(bytes, { projection: "content", videoProjection: "iso-bmff-v1" });
    try {
      assert.deepEqual(s.document.slides[0].elements.filter(x => x.type === "video").map(x => [x.playbackVolume, x.muted, x.loop, x.rewindAfterPlayback, x.playbackStart]), expected);
      bytes = (await generatePptx(s.document, { videoMode: "iso-bmff-v1" })).bytes;
    } finally { s.dispose(); }
  }
  const invalid = modify(bytes, s => s.replace('vol="20000" mute="1"', 'vol="12345" mute="1"'));
  const s = await importPptx(invalid, { projection: "content", videoProjection: "iso-bmff-v1" });
  try { const v = s.document.slides[0].elements.find(x => x.type === "video"); assert.equal(v.playbackVolume, undefined); assert.equal(v.playbackStart, undefined); } finally { s.dispose(); }
});
