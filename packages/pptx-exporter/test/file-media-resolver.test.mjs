import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile, symlink, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { constants } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createPresentation } from '@pptkit/core';
import { createFileMediaResolver, generatePptx } from '../dist/node.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

test('authorized file media is bounded, shared and retryable', { skip: !constants.O_NOFOLLOW || !constants.O_NONBLOCK }, async t => {
  const dir = await mkdtemp(join(tmpdir(), 'pptkit-media-')); t.after(() => rm(dir, { recursive: true, force: true }));
  const wave = readGeneratedZip(await readFile(new URL('../../../tests/fixtures/pcm-wave-powerpoint.pptx', import.meta.url))).get('ppt/media/media1.wav');
  const path = join(dir, 'tone.wav'), link = join(dir, 'link.wav'), missing = join(dir, 'missing.wav');
  await writeFile(path, wave); await symlink(path, link);
  const make = value => {
    const p = createPresentation(); p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'path', value } });
    p.addSlide({ elements: [0, 1].map(i => ({ type: 'audio', id: `a${i}`, audioAssetId: 'sound', box: { x: i * 50, y: 0, width: 40, height: 40 } })) }); return p;
  };
  const calls = [], resolver = createFileMediaResolver({ authorizePath: value => { calls.push(value); return value !== missing; } });
  const options = { audioMode: 'wave-pcm-v1', resolveMedia: resolver, maxAssetBytes: wave.length };
  const result = await generatePptx(make(join(dir, '.', 'tone.wav')), options);
  assert.deepEqual(calls, [path]); assert.deepEqual(readGeneratedZip(result.bytes).get('ppt/media/sound.wav'), wave);
  await assert.rejects(generatePptx(make(missing), options), /not authorized/);
  await assert.rejects(generatePptx(make(dir), options), /regular file/);
  await assert.rejects(generatePptx(make(link), options));
  await assert.rejects(generatePptx(make(path), { ...options, maxAssetBytes: wave.length - 1 }), /maxBytes/);
  await writeFile(path, Uint8Array.of(1, 2, 3)); await assert.rejects(generatePptx(make(path), options));
  await writeFile(path, wave); assert.deepEqual(readGeneratedZip((await generatePptx(make(path), options)).bytes).get('ppt/media/sound.wav'), wave);
});

test('file resolver rejects wrong sources and cancels pending authorization', async () => {
  assert.throws(() => createFileMediaResolver({}), /authorizePath/);
  let calls = 0; const resolver = createFileMediaResolver({ authorizePath() { calls++; return new Promise(() => {}); } });
  await assert.rejects(resolver({ source: { type: 'url', value: 'https://example.test/a' } }, { maxBytes: 10 }), /absolute path/);
  await assert.rejects(resolver({ source: { type: 'path', value: 'relative.wav' } }, { maxBytes: 10 }), /absolute path/);
  await assert.rejects(resolver({ source: { type: 'path', value: '/tmp/a' } }, { maxBytes: 0 }), /maxBytes/);
  if (!constants.O_NOFOLLOW || !constants.O_NONBLOCK) return;
  const controller = new AbortController(), reason = new Error('cancel authorization');
  const pending = resolver({ source: { type: 'path', value: '/tmp/a' } }, { maxBytes: 10, signal: controller.signal });
  await new Promise(resolve => setTimeout(resolve, 0)); controller.abort(reason);
  await assert.rejects(pending, error => error === reason); assert.equal(calls, 1);
});

test('the file adapter supports video and leaves unreferenced assets untouched', { skip: !constants.O_NOFOLLOW || !constants.O_NONBLOCK }, async t => {
  const dir = await mkdtemp(join(tmpdir(), 'pptkit-video-file-')); t.after(() => rm(dir, { recursive: true, force: true }));
  const video = readGeneratedZip(await readFile(new URL('../../../tests/fixtures/h264-mp4-powerpoint.pptx', import.meta.url))).get('ppt/media/media1.mp4');
  const path = join(dir, 'movie.mp4'); await writeFile(path, video);
  const p = createPresentation();
  p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'path', value: path } });
  p.registerAsset({ id: 'unused', kind: 'video', mimeType: 'video/mp4', source: { type: 'path', value: join(dir, 'missing') } });
  p.addSlide({ elements: [{ type: 'video', videoAssetId: 'movie', box: { x: 0, y: 0, width: 120, height: 60 } }] });
  const calls = [], result = await generatePptx(p, { videoMode: 'iso-bmff-v1', resolveMedia: createFileMediaResolver({ authorizePath(value) { calls.push(value); return true; } }) });
  assert.deepEqual(calls, [path]); assert.deepEqual(readGeneratedZip(result.bytes).get('ppt/media/movie.mp4'), video);
});

test('a FIFO is rejected without waiting for a writer', { skip: !constants.O_NOFOLLOW || !constants.O_NONBLOCK || process.platform === 'win32', timeout: 3000 }, async t => {
  const dir = await mkdtemp(join(tmpdir(), 'pptkit-media-fifo-')); t.after(() => rm(dir, { recursive: true, force: true }));
  const path = join(dir, 'pipe'); execFileSync('mkfifo', [path]);
  const resolver = createFileMediaResolver({ authorizePath: () => true });
  await assert.rejects(resolver({ source: { type: 'path', value: path } }, { maxBytes: 1024 }), /regular file/);
});
