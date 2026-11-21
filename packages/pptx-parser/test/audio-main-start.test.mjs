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
