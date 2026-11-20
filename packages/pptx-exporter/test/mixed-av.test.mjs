import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, mkdtemp, rm } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';
import { importPptx } from '../../pptx-parser/dist/index.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const fixture = new URL('../../../tests/fixtures/pptkit-mixed-av.pptx', import.meta.url);
const options = { videoMode: 'iso-bmff-v1', audioMode: 'wave-pcm-v1' };

test('mixed sample is reproducible and survives import and fresh export', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'pptkit-mixed-av-'));
  try {
    const out = join(dir, 'mixed.pptx');
    const result = spawnSync(process.execPath, [fileURLToPath(new URL('../../../scripts/research/generate-mixed-av.mjs', import.meta.url)), out], { encoding: 'utf8' });
    assert.equal(result.status, 0, result.stderr);
    const bytes = await readFile(fixture); assert.deepEqual(await readFile(out), bytes);
    const session = await importPptx(bytes, { projection: 'content', videoProjection: 'iso-bmff-v1', audioProjection: 'wave-pcm-v1' });
    try {
      const elements = session.document.slides[0].elements.filter(e => e.type === 'audio' || e.type === 'video');
      assert.deepEqual(elements.map(e => e.type), ['video', 'video', 'audio', 'audio']);
      assert.equal(elements[0].videoAssetId, elements[1].videoAssetId);
      assert.equal(elements[2].audioAssetId, elements[3].audioAssetId);
      const plan = await session.prepareSave({ mode: 'no-op' });
      assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
      const fresh = await generatePptx(session.document, options);
      const again = await importPptx(fresh.bytes, { projection: 'content', videoProjection: 'iso-bmff-v1', audioProjection: 'wave-pcm-v1' });
      try { assert.deepEqual(again.document.slides[0].elements.filter(e => e.type === 'audio' || e.type === 'video').map(e => e.type), ['video', 'video', 'audio', 'audio']); }
      finally { again.dispose(); }
    } finally { session.dispose(); }
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('interleaved audio and video keep object order and independent timing', async () => {
  const source = readGeneratedZip(await readFile(fixture)), p = createPresentation();
  p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: source.get('ppt/media/movie.mp4') } });
  p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/x-wav', source: { type: 'bytes', value: source.get('ppt/media/sound.wav') } });
  p.addSlide({ elements: ['audio', 'video', 'audio', 'video'].map((type, i) => ({ type, id: `m${i}`, ...(type === 'audio' ? { audioAssetId: 'sound' } : { videoAssetId: 'movie' }), box: { x: i * 50, y: 0, width: 40, height: 40 } })) });
  const output = await generatePptx(p, options), parts = readGeneratedZip(output.bytes), xml = parts.get('ppt/slides/slide1.xml').toString();
  const ids = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map(m => m[1]); assert.equal(ids.length, new Set(ids).size);
  const commands = [...xml.matchAll(/<p:cmd type="call" cmd="([^"]+)"><p:cBhvr><p:cTn id="\d+" dur="(\d+)" fill="hold"\/><p:tgtEl><p:spTgt spid="(\d+)"/g)].map(m => [m[1], m[2], m[3]]);
  assert.deepEqual(commands, [['playFrom(0.0)', '1000', '2'], ['playFrom(0.0)', '2000', '3'], ['playFrom(0.0)', '1000', '4'], ['playFrom(0.0)', '2000', '5'], ['togglePause', '1', '3'], ['togglePause', '1', '5']]);
  for (const missing of [{ videoMode: options.videoMode }, { audioMode: options.audioMode }]) await assert.rejects(generatePptx(p, missing), /requires (audioMode|videoMode)/);
});
