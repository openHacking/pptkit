import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { importPptx } from "../dist/index.js";
import { generatePptx } from "../../pptx-exporter/dist/index.js";
import { picture, imageFixture, encodeImages } from "./image-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const options = { projection: "content", audioProjection: "wave-pcm-v1" };
const sample = name => new URL(`../../../tests/fixtures/${name}.pptx`, import.meta.url);
const trims = s => s.document.slides[0].elements.filter(x => x.type === "audio").map(x => [x.trimStartMs, x.trimEndMs]);

test("native and generated audio trims project per instance and survive fresh export", async () => {
  for (const [name, expected] of [["powerpoint-audio-trim", [[undefined, undefined], [undefined, undefined], [100, 50], [undefined, undefined]]], ["pptkit-audio-trim", [[undefined, undefined], [100, 0], [100, 50], [0, 200]]]]) {
    const bytes = await readFile(sample(name)), original = readGeneratedZip(bytes), wave = [...original].find(([part]) => part.endsWith(".wav"))[1], session = await importPptx(bytes, options);
    try {
      assert.deepEqual(trims(session), expected); assert.equal(session.document.assets.filter(x => x.kind === "audio").length, 1);
      const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
      assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
      let fresh = (await generatePptx(session.document, { audioMode: "wave-pcm-v1" })).bytes;
      for (let round = 0; round < 3; round++) {
        const reopened = await importPptx(fresh, options);
        try { assert.deepEqual(trims(reopened), expected); const source = readGeneratedZip(fresh); assert.equal([...source.keys()].filter(x => x.endsWith(".wav")).length, 1); assert.deepEqual([...source].find(([part]) => part.endsWith(".wav"))[1], wave); fresh = (await generatePptx(reopened.document, { audioMode: "wave-pcm-v1" })).bytes; }
        finally { reopened.dispose(); }
      }
      assert.ok(session.report.diagnostics.some(x => x.code === "media-playback-unmapped"));
    } finally { session.dispose(); }
  }
});

function fixture({ direct = "NULL", trim = '<m:trim st="2" end="3"/>', office = "../media/tone.wav", extra = "" } = {}) {
  const pic = picture({ id: 2 }).replace("<p:nvPr/>", `<p:nvPr><a:audioFile r:link="sound"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><m:media xmlns:m="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="office">${trim}</m:media>${extra}</p:ext></p:extLst></p:nvPr>`);
  const f = imageFixture(pic), wave = Buffer.alloc(124, 128); wave.write("RIFF", 0); wave.writeUInt32LE(116, 4); wave.write("WAVEfmt ", 8); wave.writeUInt32LE(16, 16); wave.writeUInt16LE(1, 20); wave.writeUInt16LE(1, 22); wave.writeUInt32LE(8000, 24); wave.writeUInt32LE(8000, 28); wave.writeUInt16LE(1, 32); wave.writeUInt16LE(8, 34); wave.write("data", 36); wave.writeUInt32LE(80, 40);
  f.parts["media/tone.wav"] = wave; f.mediaTypes["media/tone.wav"] = "audio/wav";
  const rel = f.rel("sound", "audio", direct === "internal" ? "../media/tone.wav" : direct);
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin"), ...(direct === "internal" ? [rel] : [{ ...rel, TargetMode: "External" }]), { Id: "office", Type: "http://schemas.microsoft.com/office/2007/relationships/media", Target: office }]);
  return encodeImages(f);
}

test("native NULL fallback requires one valid bounded embedded trim and no ambiguous references", async () => {
  const good = await importPptx(fixture(), options);
  try { assert.deepEqual(trims(good), [[2, 3]]); } finally { good.dispose(); }
  for (const opts of [
    { direct: "https://example.invalid/sound.wav" }, { direct: "null" }, { trim: "" },
    { trim: '<m:trim st="-1" end="0"/>' }, { trim: '<m:trim st="1.5" end="0"/>' },
    { trim: '<m:trim st="7" end="3"/>' }, { trim: '<m:trim st="2" end="3"/><m:trim st="1"/>' },
    { trim: '<m:trim st="2" extra="1"/>' }, { trim: '<m:trim st="2">unexpected</m:trim>' },
    { extra: '<m:media xmlns:m="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="office"/>' },
  ]) {
    const bytes = fixture(opts), session = await importPptx(bytes, options);
    try { assert.equal(trims(session).length, 0, JSON.stringify(opts)); assert.ok(session.report.findings.some(x => x.feature === "drawing.audio.projected-fields" && x.tier === "unsupported")); const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes)); }
    finally { session.dispose(); }
  }
});

test("unsupported trim on an otherwise valid internal source is diagnosed without fabrication", async () => {
  for (const trim of ['<m:trim st="1.5"/>', '<m:trim st="9" end="1"/>', '<m:trim st="2"/><m:trim st="3"/>']) {
    const session = await importPptx(fixture({ direct: "internal", trim }), options);
    try { assert.deepEqual(trims(session), [[undefined, undefined]]); assert.ok(session.report.findings.some(x => x.feature === "drawing.audio.trim" && x.tier === "unsupported")); }
    finally { session.dispose(); }
  }
});
