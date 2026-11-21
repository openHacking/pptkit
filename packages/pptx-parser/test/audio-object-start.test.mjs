import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { importPptx } from "../dist/index.js";
import { generatePptx } from "../../pptx-exporter/dist/index.js";
const { zipSync, unzipSync, strFromU8, strToU8 } = createRequire(import.meta.url)("fflate");
const options = { projection: "content", audioProjection: "wave-pcm-v1" };
const sample = name => readFile(new URL(`../../../tests/fixtures/${name}.pptx`, import.meta.url));
const starts = s => s.document.slides[0].elements.filter(x => x.type === "audio").map(x => x.playbackStart);

test("native and generated independent clicks project by source identity", async () => {
  for (const [name, expected] of [["powerpoint-single-audio-object-click", ["object-click"]], ["pptkit-single-audio-object-click", ["object-click"]], ["powerpoint-audio-start-object-click", [undefined, "object-click"]], ["pptkit-audio-start-object-click", [undefined, "object-click"]], ["pptkit-audio-trim", [undefined, "object-click", undefined, "object-click"]]]) {
    const bytes = await sample(name), s = await importPptx(bytes, options);
    try { assert.deepEqual(starts(s), expected, name); const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes)); assert.ok(s.report.diagnostics.some(x => x.code === "media-playback-unmapped")); }
    finally { s.dispose(); }
  }
  let bytes = await sample("pptkit-single-audio-object-click");
  for (let i = 0; i < 3; i++) { const s = await importPptx(bytes, options); try { assert.deepEqual(starts(s), ["object-click"]); bytes = (await generatePptx(s.document, { audioMode: "wave-pcm-v1" })).bytes; } finally { s.dispose(); } }
});

function modify(bytes, mutate) {
  const entries = unzipSync(bytes), name = "ppt/slides/slide1.xml", input = strFromU8(entries[name]), output = mutate(input); assert.notEqual(output, input); entries[name] = strToU8(output); return zipSync(entries);
}
test("independent click mapping rejects malformed or conflicting control chains", async () => {
  const original = await sample("pptkit-single-audio-object-click");
  const changes = [
    x => x.replace('evt="onClick" delay="0"', 'evt="onMouseOver" delay="0"'),
    x => x.replace('cmd="playFrom(0.0)"', 'cmd="togglePause"'),
    x => x.replace('dur="1000" fill="hold"', 'dur="999" fill="hold"'),
    x => x.replace('nodeType="clickEffect"', 'nodeType="withEffect"'),
    x => x.replace('evtFilter="cancelBubble"', 'evtFilter="other"'),
    x => x.replace('<p:spTgt spid="2"/>', '<p:spTgt spid="3"/>'),
    x => x.replace('<p:cTn id="6" dur=', '<p:cTn id="02" dur='),
    x => x.replace('</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>', '<p:cmd type="call" cmd="togglePause"><p:cBhvr><p:cTn id="99"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'),
  ];
  for (const change of changes) {
    const bytes = modify(original, change), s = await importPptx(bytes, options);
    try { assert.deepEqual(starts(s), [undefined]); assert.ok(s.report.diagnostics.some(x => x.code === "media-playback-unmapped")); const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, bytes); }
    finally { s.dispose(); }
  }
});
