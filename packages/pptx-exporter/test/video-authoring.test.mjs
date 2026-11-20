import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';
import { importPptx } from '../../pptx-parser/dist/index.js';
import { readFile } from 'node:fs/promises';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

async function sample() {
  const bytes = await readFile(new URL('../../../tests/fixtures/h264-mp4-powerpoint.pptx', import.meta.url));
  return readGeneratedZip(bytes).get('ppt/media/media1.mp4');
}

test('generated native-observation fixture remains reproducible', async () => {
  const source = await readFile(new URL('../../../tests/fixtures/h264-mp4-powerpoint.pptx', import.meta.url));
  const expected = await readFile(new URL('../../../tests/fixtures/h264-mp4-pptkit-generated.pptx', import.meta.url));
  const s = await importPptx(source, { projection: 'content', videoProjection: 'iso-bmff-v1' });
  try { assert.deepEqual(Buffer.from((await generatePptx(s.document, { videoMode: 'iso-bmff-v1' })).bytes), expected); }
  finally { s.dispose(); }
});

test('same-slide video supports independent timing targets and shared or distinct assets', async () => {
  const movie = await sample(), short = Buffer.from(movie), mvhd = short.indexOf(Buffer.from('mvhd'));
  short.writeUInt32BE(1000, mvhd + 20); // A distinct metadata fixture, not a validated shorter clip.
  const p = createPresentation();
  for (const [id, bytes] of [['shared', movie], ['other', short]]) p.registerAsset({ id, kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: bytes } });
  p.addSlide({ elements: ['shared', 'shared', 'other'].map((videoAssetId, i) => ({ type: 'video', id: `v${i}`, videoAssetId, box: { x: i * 130, y: 20, width: 120, height: 60 } })) });
  const output = await generatePptx(p, { videoMode: 'iso-bmff-v1' }), parts = readGeneratedZip(output.bytes), xml = parts.get('ppt/slides/slide1.xml').toString();
  assert.deepEqual([...parts.keys()].filter(name => name.endsWith('.mp4')).sort(), ['ppt/media/other.mp4', 'ppt/media/shared.mp4']);
  const ids = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map(match => match[1]); assert.equal(ids.length, new Set(ids).size);
  const commands = [...xml.matchAll(/<p:cmd type="call" cmd="([^"]+)"><p:cBhvr><p:cTn id="\d+" dur="(\d+)" fill="hold"\/><p:tgtEl><p:spTgt spid="(\d+)"/g)].map(m => [m[1], m[2], m[3]]);
  assert.deepEqual(commands, [['playFrom(0.0)', '2000', '2'], ['playFrom(0.0)', '2000', '3'], ['playFrom(0.0)', '1000', '4'], ['togglePause', '1', '2'], ['togglePause', '1', '3'], ['togglePause', '1', '4']]);
  const imported = await importPptx(output.bytes, { projection: 'content', videoProjection: 'iso-bmff-v1' });
  try {
    const elements = imported.document.slides[0].elements.filter(x => x.type === 'video');
    assert.equal(elements.length, 3); assert.equal(elements[0].videoAssetId, elements[1].videoAssetId); assert.notEqual(elements[0].videoAssetId, elements[2].videoAssetId);
    const plan = await imported.prepareSave({ mode: 'no-op' }); assert.deepEqual((await imported.save(plan.plan)).bytes, output.bytes);
    const second = await generatePptx(imported.document, { videoMode: 'iso-bmff-v1' });
    const again = await importPptx(second.bytes, { projection: 'content', videoProjection: 'iso-bmff-v1' });
    try { assert.equal(again.document.slides[0].elements.filter(x => x.type === 'video').length, 3); } finally { again.dispose(); }
  } finally { imported.dispose(); }
});

test('retained two-video sample is reproducible', async () => {
  const { spawnSync } = await import('node:child_process');
  const { mkdtemp, readFile, rm } = await import('node:fs/promises');
  const { fileURLToPath } = await import('node:url');
  const { tmpdir } = await import('node:os'), { join } = await import('node:path');
  const dir = await mkdtemp(join(tmpdir(), 'pptkit-video-'));
  try {
    const output = join(dir, 'multi.pptx'), script = new URL('../../../scripts/research/generate-multi-video.mjs', import.meta.url);
    const result = spawnSync(process.execPath, [fileURLToPath(script), output], { encoding: 'utf8' }); assert.equal(result.status, 0, result.stderr);
    assert.deepEqual(await readFile(output), await readFile(new URL('../../../tests/fixtures/h264-mp4-pptkit-multi.pptx', import.meta.url)));
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('PPTX rejects video before resolving any assets, including hidden nested video', async () => {
  for (const nested of [false, true]) {
    const p = createPresentation();
    p.registerAsset({ id: 'video', kind: 'video', source: { type: 'url', value: 'https://example.invalid/video.mp4' } });
    p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'url', value: 'https://example.invalid/poster.png' } });
    const box = { x: 0, y: 0, width: 120, height: 60 }, video = { type: 'video', hidden: nested, videoAssetId: 'video', posterAssetId: 'poster', box };
    p.addSlide({ elements: nested ? [{ type: 'group', box, coordinateSize: { width: 120, height: 60 }, children: [video] }] : [video] });
    let calls = 0;
    await assert.rejects(generatePptx(p, { resolveAsset: async () => { calls++; throw Error('unexpected'); } }), /requires videoMode iso-bmff-v1/);
    assert.equal(calls, 0);
  }
});

test('unreferenced video assets are not passed to the image loader', async () => {
  const p = createPresentation();
  p.registerAsset({ id: 'unused', kind: 'video', source: { type: 'url', value: 'https://example.invalid/video.mp4' } });
  p.addSlide({ elements: [] }); let calls = 0;
  await generatePptx(p, { resolveAsset: async () => { calls++; throw Error('unexpected'); } });
  assert.equal(calls, 0);
});

test('explicit MP4 export preserves video bytes, poster and native timing across slides', async () => {
  const mp4 = await sample(), p = createPresentation();
  p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: mp4 } });
  for (const id of ['one', 'two']) p.addSlide({ id, elements: [{ type: 'video', id: `video-${id}`, videoAssetId: 'movie', box: { x: 10, y: 20, width: 120, height: 60 }, transform: { rotation: 90 } }] });
  const output = await generatePptx(p, { videoMode: 'iso-bmff-v1' }), parts = readGeneratedZip(output.bytes);
  assert.deepEqual(parts.get('ppt/media/movie.mp4'), mp4); assert.ok(parts.has('ppt/media/pptkit-video-poster.png'));
  assert.equal([...parts.keys()].filter(name => name.endsWith('.mp4')).length, 1);
  assert.equal(output.warnings.filter(w => w.code === 'video-playback-unverified').length, 2);
  for (const index of [1, 2]) {
    const xml = parts.get(`ppt/slides/slide${index}.xml`).toString(), rel = parts.get(`ppt/slides/_rels/slide${index}.xml.rels`).toString();
    assert.match(xml, /<a:videoFile r:link="rId\d+"\/>/); assert.match(xml, /<p:video>/); assert.match(xml, /cmd="togglePause"/); assert.match(xml, /dur="2000"/); assert.match(xml, /rot="5400000"/);
    assert.match(rel, /relationships\/video/); assert.match(rel, /2007\/relationships\/media/); assert.match(rel, /relationships\/image/);
  }
  const s = await importPptx(output.bytes, { projection: 'content', videoProjection: 'iso-bmff-v1' });
  try {
    const video = s.document.slides.map(slide => slide.elements.find(x => x.type === 'video'));
    assert.equal(video[0].videoAssetId, video[1].videoAssetId); assert.ok(s.report.videoSources.filter(x => x.companionStatus === 'same-part').length === 2);
    assert.equal(s.document.assets.find(x => x.id === video[0].videoAssetId).source.value, Buffer.from(mp4).toString('base64'));
  } finally { s.dispose(); }
});

test('authored video posters must resolve and never trigger video fetching', async () => {
  const p = createPresentation(); p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: await sample() } });
  p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'url', value: 'https://example.invalid/poster.png' } });
  p.addSlide({ elements: [{ type: 'video', videoAssetId: 'movie', posterAssetId: 'poster', box: { x: 0, y: 0, width: 120, height: 60 } }] });
  const calls = [];
  await assert.rejects(generatePptx(p, { videoMode: 'iso-bmff-v1', resolveAsset: async asset => { calls.push(asset.id); return undefined; } }), /video or its poster could not be packaged/);
  assert.deepEqual(calls, ['poster']);
});

test('unsafe or unsupported MP4 exports fail before resolving any poster', async () => {
  const original = await sample(), mvhd = original.indexOf(Buffer.from('mvhd')); assert.ok(mvhd > 0);
  for (const mode of ['wrong-mime', 'url', 'truncated', 'missing-movie', 'missing-data', 'fragmented', 'short-mvhd', 'version', 'zero-scale', 'zero-duration', 'unknown-duration', 'duplicate-mvhd', 'nested', 'hidden', 'too-large', 'mixed-audio']) {
    let bytes = Buffer.from(original);
    if (mode === 'truncated') bytes = bytes.subarray(0, bytes.length - 1);
    if (mode === 'missing-movie') bytes.write('xxxx', 36);
    if (mode === 'missing-data') bytes.write('xxxx', 1207);
    if (mode === 'fragmented') bytes.write('moof', 1199);
    if (mode === 'short-mvhd') bytes.writeUInt32BE(28, mvhd - 4);
    if (mode === 'version') bytes[mvhd + 4] = 1;
    if (mode === 'zero-scale') bytes.writeUInt32BE(0, mvhd + 16);
    if (mode === 'zero-duration') bytes.writeUInt32BE(0, mvhd + 20);
    if (mode === 'unknown-duration') bytes.writeUInt32BE(0xffffffff, mvhd + 20);
    if (mode === 'duplicate-mvhd') bytes.write('mvhd', bytes.indexOf(Buffer.from('trak')));
    const p = createPresentation();
    p.registerAsset({ id: 'video', kind: 'video', mimeType: mode === 'wrong-mime' ? 'video/other' : 'video/mp4', source: mode === 'url' ? { type: 'url', value: 'https://example.invalid/movie.mp4' } : { type: 'bytes', value: bytes } });
    p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'url', value: 'https://example.invalid/poster.png' } });
    const box = { x: 0, y: 0, width: 120, height: 60 }, video = { type: 'video', hidden: mode === 'hidden', videoAssetId: 'video', posterAssetId: 'poster', box };
    let elements = mode === 'nested' ? [{ type: 'group', box, coordinateSize: { width: 120, height: 60 }, children: [video] }] : [video];
    if (mode === 'mixed-audio') { p.registerAsset({ id: 'sound', kind: 'audio', source: { type: 'bytes', value: Uint8Array.of(1) } }); elements = [video, { type: 'audio', audioAssetId: 'sound', box }]; }
    p.addSlide({ elements }); let calls = 0;
    await assert.rejects(generatePptx(p, { videoMode: 'iso-bmff-v1', ...(mode === 'too-large' ? { maxAssetBytes: original.length - 1 } : {}), resolveAsset: async () => { calls++; throw Error('unexpected'); } }), undefined, mode);
    assert.equal(calls, 0, mode);
  }
});

test('same-slide video count is bounded before resolving any poster', async () => {
  const p = createPresentation(); p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: await sample() } });
  p.addSlide({ elements: Array.from({ length: 4097 }, () => ({ type: 'video', videoAssetId: 'movie', box: { x: 0, y: 0, width: 10, height: 10 } })) });
  let calls = 0; await assert.rejects(generatePptx(p, { videoMode: 'iso-bmff-v1', resolveAsset: async () => { calls++; throw Error('unexpected'); } }), /4096 video objects/);
  assert.equal(calls, 0);
});
