import test from "node:test";
import assert from "node:assert/strict";
import { importPptx } from "../dist/index.js";
import { picture, imageFixture, encodeImages } from "./image-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { importPptxFile } from "../dist/node.js";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

function fixture(strict = false) {
  const media = id => picture({ id }).replace('<p:nvPr/>', '<p:nvPr><a:audioFile r:link="sound"/></p:nvPr>');
  const f = imageFixture(media(2) + media(3), strict), wave = Buffer.alloc(124, 128);
  wave.write('RIFF', 0); wave.writeUInt32LE(116, 4); wave.write('WAVEfmt ', 8); wave.writeUInt32LE(16, 16);
  wave.writeUInt16LE(1, 20); wave.writeUInt16LE(1, 22); wave.writeUInt32LE(8000, 24); wave.writeUInt32LE(8000, 28);
  wave.writeUInt16LE(1, 32); wave.writeUInt16LE(8, 34); wave.write('data', 36); wave.writeUInt32LE(80, 40);
  f.parts['media/tone.wav'] = wave; f.mediaTypes['media/tone.wav'] = 'audio/wav';
  f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), f.rel('sound', 'audio', '../media/tone.wav')]);
  return f;
}
const options = { projection: 'content', audioSourceProfile: 'wave-pcm-v1' };

test('audio source inventory identifies shared PCM and preserves original bytes', async () => {
  for (const strict of [false, true]) {
    const bytes = encodeImages(fixture(strict)), session = await importPptx(bytes, options);
    try {
      assert.equal(session.report.audioSources.length, 2);
      assert.equal(session.document.assets.length, 1); // Only the picture poster is projected.
      for (const [index, item] of session.report.audioSources.entries()) {
        assert.equal(item.status, 'wave-pcm'); assert.equal(item.mediaPart, '/media/tone.wav');
        assert.equal(item.relationshipId, 'sound'); assert.equal(item.source.sourceObjectId, String(index + 2));
        assert.match(item.source.path, /audioFile\[1\]$/); assert.ok(item.sourceElementIndex > 0);
        assert.deepEqual(item.pcm, { channels: 1, sampleRate: 8000, bitsPerSample: 8, frameCount: 80, durationSeconds: .01 });
        assert.ok(Object.isFrozen(item.pcm));
      }
      const plan = await session.prepareSave({ mode: 'no-op' });
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});

test('file adapter captures the opt-in source profile before I/O', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'pptkit-audio-'));
  try {
    const path = join(dir, 'source.pptx'); await writeFile(path, encodeImages(fixture()));
    const config = { ...options }, pending = importPptxFile(path, config); config.audioSourceProfile = undefined;
    const { session } = await pending;
    try { assert.equal(session.report.audioSources.length, 2); } finally { session.dispose(); }
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('audio inventory reports external and unsupported audio without fetching or guessing', async () => {
  for (const mode of ['external', 'missing', 'type', 'damaged', 'mime']) {
    const f = fixture();
    if (mode === 'external') f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), { ...f.rel('sound', 'audio', 'https://example.invalid/tone.wav'), TargetMode: 'External' }]);
    if (mode === 'missing') f.parts['deck/z.xml'] = f.parts['deck/z.xml'].replaceAll('r:link="sound"', 'r:link="unknown"');
    if (mode === 'type') f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), f.rel('sound', 'image', '../media/tone.wav')]);
    if (mode === 'damaged') f.parts['media/tone.wav'][0] = 0;
    if (mode === 'mime') f.mediaTypes['media/tone.wav'] = 'audio/mpeg';
    const session = await importPptx(encodeImages(f), options);
    try { assert.deepEqual(session.report.audioSources.map(x => x.status), [mode === 'external' ? 'external' : 'unsupported', mode === 'external' ? 'external' : 'unsupported']); }
    finally { session.dispose(); }
  }
  await assert.rejects(importPptx(encodeImages(fixture()), { ...options, sessionLimits: { maxImportedAssetBytes: 123 } }), { code: 'import-limit-exceeded' });
  const shared = await importPptx(encodeImages(fixture()), { ...options, sessionLimits: { maxImportedAssetBytes: 124, maxTotalImportedAssetBytes: 124 } });
  assert.equal(shared.report.audioSources.length, 2); shared.dispose();
  const ordinary = await importPptx(encodeImages(fixture()), { projection: 'content' });
  assert.equal(ordinary.report.audioSources, undefined); ordinary.dispose();
  await assert.rejects(importPptx(encodeImages(fixture()), { ...options, audioSourceProfile: 'other' }), TypeError);
  await assert.rejects(importPptx(encodeImages(fixture()), { ...options, projection: 'structure' }), TypeError);
});
