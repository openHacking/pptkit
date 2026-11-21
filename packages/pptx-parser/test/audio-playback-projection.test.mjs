import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { importPptx } from "../dist/index.js";
import { generatePptx } from "../../pptx-exporter/dist/index.js";
import { picture, imageFixture, encodeImages } from "./image-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
const options = { projection: "content", audioProjection: "wave-pcm-v1" };
const values = s => s.document.slides[0].elements.filter(x => x.type === "audio").map(x => [x.playbackVolume, x.loop, x.rewindAfterPlayback]);

test("native and generated audio settings stay per object across fresh exports", async () => {
  const expected = [["low", false, false], ["medium", false, false], ["high", false, false], ["mute", false, false]];
  for (const [name, selected] of [["pptkit-audio-volume", expected], ["powerpoint-audio-loop", expected.map((v, i) => i === 2 ? ["high", true, false] : v)], ["powerpoint-audio-rewind", expected.map((v, i) => i === 2 ? ["high", false, true] : v)], ["powerpoint-audio-loop-rewind", expected.map((v, i) => i === 2 ? ["high", true, true] : v)]]) {
    const bytes = await readFile(new URL(`../../../tests/fixtures/${name}.pptx`, import.meta.url)), s = await importPptx(bytes, options);
    try {
      assert.deepEqual(values(s), selected); assert.equal(s.document.assets.filter(x => x.kind === "audio").length, 1);
      const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
      let exported = (await generatePptx(s.document, { audioMode: "wave-pcm-v1" })).bytes;
      for (let i = 0; i < 3; i++) { const reopened = await importPptx(exported, options); try { assert.deepEqual(values(reopened), selected); exported = (await generatePptx(reopened.document, { audioMode: "wave-pcm-v1" })).bytes; } finally { reopened.dispose(); } }
    } finally { s.dispose(); }
  }
});

function fixture(nodes) {
  const f = imageFixture(picture({ id: 2 }).replace("<p:nvPr/>", '<p:nvPr><a:audioFile r:link="sound"/></p:nvPr>'));
  const wave = Buffer.alloc(124, 128); wave.write("RIFF", 0); wave.writeUInt32LE(116, 4); wave.write("WAVEfmt ", 8); wave.writeUInt32LE(16, 16); wave.writeUInt16LE(1, 20); wave.writeUInt16LE(1, 22); wave.writeUInt32LE(8000, 24); wave.writeUInt32LE(8000, 28); wave.writeUInt16LE(1, 32); wave.writeUInt16LE(8, 34); wave.write("data", 36); wave.writeUInt32LE(80, 40);
  f.parts["media/tone.wav"] = wave; f.mediaTypes["media/tone.wav"] = "audio/wav";
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin"), f.rel("sound", "audio", "../media/tone.wav")]);
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("</p:sld>", `<p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>${nodes}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing></p:sld>`);
  return encodeImages(f);
}
function media({ volume = 'vol="20000"', time = 'id="2" fill="hold" display="0"', target = '<p:spTgt spid="2"/>', start = '<p:cond delay="indefinite"/>', end = '<p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>' } = {}) {
  return `<p:audio><p:cMediaNode ${volume}><p:cTn ${time}><p:stCondLst>${start}</p:stCondLst><p:endCondLst>${end}</p:endCondLst></p:cTn><p:tgtEl>${target}</p:tgtEl></p:cMediaNode></p:audio>`;
}

test("unrecognized or conflicting media settings never fabricate values", async () => {
  const valid = await importPptx(fixture(media()), options);
  try { assert.deepEqual(values(valid), [["low", false, false]]); } finally { valid.dispose(); }
  for (const node of [media() + media(), media() + '<p:par><p:cTn id="02"/></p:par>', media({ volume: 'vol="30000"' }), media({ volume: 'mute="1" vol="80000"' }), media({ time: 'id="2" fill="hold" display="0" repeatCount="2000"' }), media({ time: 'id="2" fill="hold" display="0" dur="500"' }), media({ start: '<p:cond delay="indefinite" evt="onClick"/>' }), media({ end: '<p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/><p:sldTgt/></p:tgtEl></p:cond>' }), media({ target: '<p:spTgt spid="2"><p:txEl/></p:spTgt>' })]) {
    const bytes = fixture(node), s = await importPptx(bytes, options);
    try { assert.deepEqual(values(s), [[undefined, undefined, undefined]]); assert.ok(s.report.findings.some(x => x.feature === "drawing.audio.playback" && x.tier === "unsupported")); const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes)); }
    finally { s.dispose(); }
  }
});
