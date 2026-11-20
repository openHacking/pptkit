import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createPresentation, normalizePresentation, validatePresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';
import { importPptx } from '../../pptx-parser/dist/index.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const root = fileURLToPath(new URL('../../../', import.meta.url));
const source = new URL('../../../tests/fixtures/pptkit-audio-start-object-click.pptx', import.meta.url);

test('generated mixed object-click sample is reproducible and matches native structure', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'pptkit-audio-click-'));
  try {
    const output = join(dir, 'sample.pptx');
    execFileSync(process.execPath, [join(root, 'scripts/research/generate-object-click-media.mjs'), output], { cwd: root });
    assert.deepEqual(await readFile(output), await readFile(source));
    execFileSync(process.execPath, [join(root, 'scripts/research/generate-object-click-media.mjs'), output, '--single'], { cwd: root });
    assert.deepEqual(await readFile(output), await readFile(join(root, 'tests/fixtures/pptkit-single-audio-object-click.pptx')));
    execFileSync('python3', [join(root, 'scripts/research/test_media_start.py')], { cwd: root, env: { ...process.env, PYTHONDONTWRITEBYTECODE: '1' } });
    execFileSync('python3', [join(root, 'scripts/research/verify-object-click-audio.py'), join(root, 'tests/fixtures/powerpoint-single-audio-object-click.pptx'), JSON.stringify({ main: [], videos: [], objects: ['4'], audio: ['4'], video: [] })], { cwd: root });
    execFileSync('python3', [join(root, 'scripts/research/verify-object-click-audio.py'), output, JSON.stringify({ main: [], videos: [], objects: ['2'], audio: ['2'], video: [] })], { cwd: root });
    const bytes = await readFile(source), session = await importPptx(bytes, { projection: 'content', audioProjection: 'wave-pcm-v1', videoProjection: 'iso-bmff-v1' });
    try {
      assert.equal(session.document.slides[0].elements.filter(x => x.type === 'audio').length, 2);
      assert.ok(session.report.diagnostics.some(x => x.code === 'media-playback-unmapped'));
      const plan = await session.prepareSave({ mode: 'no-op' }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
    } finally { session.dispose(); }
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('object-click audio coexists with sequence and video effects without timing collisions', async () => {
  const fixture = readGeneratedZip(await readFile(source)), dir = await mkdtemp(join(tmpdir(), 'pptkit-start-'));
  const movie = fixture.get('ppt/media/movie.mp4'), sound = fixture.get('ppt/media/sound.wav');
  try {
    for (const kinds of [['object'], ['object', 'object'], ['sequence', 'object', 'sequence'], ['object', 'video', 'sequence', 'video', 'object']]) {
      const p = createPresentation();
      p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: sound } });
      p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: movie } });
      p.addSlide({ elements: kinds.map((kind, i) => ({ id: `media-${i}`, ...(kind === 'video' ? { type: 'video', videoAssetId: 'movie' } : { type: 'audio', audioAssetId: 'sound', playbackStart: kind === 'object' ? 'object-click' : 'click-sequence' }), box: { x: 50 * i, y: 0, width: 40, height: 40 } })) });
      assert.deepEqual(normalizePresentation(p).slides[0].elements.filter(e => e.type === 'audio').map(e => e.playbackStart), kinds.filter(k => k !== 'video').map(k => k === 'object' ? 'object-click' : 'click-sequence'));
      const out = join(dir, 'out.pptx'); await writeFile(out, (await generatePptx(p, { audioMode: 'wave-pcm-v1', videoMode: 'iso-bmff-v1' })).bytes);
      const ids = kinds.map((_, i) => String(i + 2));
      const expected = { main: ids.filter((_, i) => kinds[i] !== 'object'), videos: ids.filter((_, i) => kinds[i] === 'video'), objects: ids.filter((_, i) => kinds[i] === 'object'), audio: ids.filter((_, i) => kinds[i] !== 'video'), video: ids.filter((_, i) => kinds[i] === 'video') };
      execFileSync('python3', [join(root, 'scripts/research/verify-object-click-audio.py'), out, JSON.stringify(expected)], { cwd: root });
    }
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('invalid audio start is diagnosed and rejected before external resolution', async () => {
  const p = createPresentation(); p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'url', value: 'https://example.invalid/a.wav' } });
  p.addSlide({ elements: [{ type: 'audio', audioAssetId: 'sound', playbackStart: 'automatic', box: { x: 0, y: 0, width: 40, height: 40 } }] });
  assert.ok(validatePresentation(p).some(x => x.code === 'invalid-audio-playback-start'));
  let calls = 0; await assert.rejects(generatePptx(p, { audioMode: 'wave-pcm-v1', resolveMedia() { calls++; } }), /Presentation validation failed/); assert.equal(calls, 0);
});

test('explicit click-sequence retains the default output bytes', async () => {
  const bytes = readGeneratedZip(await readFile(source)).get('ppt/media/sound.wav');
  const make = playbackStart => {
    const p = createPresentation(); p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: bytes } });
    p.addSlide({ elements: [{ id: 'speaker', type: 'audio', audioAssetId: 'sound', ...(playbackStart ? { playbackStart } : {}), box: { x: 0, y: 0, width: 40, height: 40 } }] }); return p;
  };
  assert.deepEqual((await generatePptx(make('click-sequence'), { audioMode: 'wave-pcm-v1' })).bytes, (await generatePptx(make(undefined), { audioMode: 'wave-pcm-v1' })).bytes);
});
