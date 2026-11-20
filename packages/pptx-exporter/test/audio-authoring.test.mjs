import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';
import { readFile } from 'node:fs/promises';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
import { importPptx } from '../../pptx-parser/dist/index.js';

const sample = async () => readGeneratedZip(await readFile(new URL('../../../tests/fixtures/pcm-wave-powerpoint.pptx', import.meta.url))).get('ppt/media/media1.wav');

test('retained PowerPoint-opened sample is reproducible from the PCM input', async () => {
  const p = createPresentation(); p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: await sample() } });
  p.addSlide({ elements: [{ type: 'audio', id: 'speaker', audioAssetId: 'sound', box: { x: 200, y: 100, width: 64, height: 64 } }] });
  const expected = await readFile(new URL('../../../tests/fixtures/pcm-wave-pptkit-generated.pptx', import.meta.url));
  assert.deepEqual(Buffer.from((await generatePptx(p, { audioMode: 'wave-pcm-v1' })).bytes), expected);
});

test('opt-in PCM WAVE output shares audio across slides and emits independent native references', async () => {
  const wave = await sample(), p = createPresentation();
  p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/x-wav', source: { type: 'bytes', value: wave } });
  for (const id of ['one', 'two']) p.addSlide({ id, elements: [{ type: 'audio', id: `audio-${id}`, audioAssetId: 'sound', box: { x: 10, y: 20, width: 64, height: 64 } }] });
  const result = await generatePptx(p, { audioMode: 'wave-pcm-v1' }), parts = readGeneratedZip(result.bytes);
  assert.deepEqual(parts.get('ppt/media/sound.wav'), wave);
  assert.ok(parts.has('ppt/media/pptkit-audio-poster.png'));
  assert.equal(result.warnings.filter(w => w.code === 'audio-playback-unverified').length, 2);
  for (const n of [1, 2]) {
    const slide = parts.get(`ppt/slides/slide${n}.xml`).toString(), rels = parts.get(`ppt/slides/_rels/slide${n}.xml.rels`).toString();
    assert.match(slide, /<a:audioFile r:link="rId\d+"\/>/); assert.match(slide, /<p14:media/);
    assert.match(slide, /playFrom\(0\.0\)/); assert.match(slide, /dur="1000"/);
    const id = /<p:cNvPr id="(\d+)" name="audio-/.exec(slide)?.[1];
    assert.ok(id); assert.deepEqual([...slide.matchAll(/<p:spTgt spid="(\d+)"/g)].map(x => x[1]), [id, id]);
    assert.match(rels, /relationships\/audio/); assert.match(rels, /2007\/relationships\/media/); assert.match(rels, /relationships\/image/);
  }
  const imported = await importPptx(result.bytes, { projection: 'content', audioSourceProfile: 'wave-pcm-v1' });
  try { assert.equal(imported.report.audioSources.length, 4); assert.equal(imported.report.audioSources.filter(x => x.companionStatus === 'same-part').length, 2); assert.ok(imported.report.audioSources.every(x => x.status === 'wave-pcm')); }
  finally { imported.dispose(); }
});

test('unsupported audio fails before external poster resolution', async () => {
  const wave = await sample();
  for (const change of ['damaged', 'wrong-mime', 'url', 'two', 'hidden', 'nested']) {
    const p = createPresentation(), bytes = Buffer.from(wave); if (change === 'damaged') bytes[0] = 0;
    p.registerAsset({ id: 'sound', kind: 'audio', mimeType: change === 'wrong-mime' ? 'audio/mpeg' : 'audio/wav', source: change === 'url' ? { type: 'url', value: 'https://example.invalid/a.wav' } : { type: 'bytes', value: bytes } });
    const audio = { type: 'audio', audioAssetId: 'sound', box: { x: 0, y: 0, width: 100, height: 50 }, ...(change === 'hidden' ? { hidden: true } : {}) };
    const elements = change === 'nested' ? [{ type: 'group', box: audio.box, coordinateSize: { width: 100, height: 50 }, children: [audio] }] : change === 'two' ? [audio, { ...audio }] : [audio];
    p.addSlide({ elements }); let calls = 0;
    await assert.rejects(generatePptx(p, { audioMode: 'wave-pcm-v1', resolveAsset: async () => { calls++; throw Error('unexpected'); } })); assert.equal(calls, 0, change);
  }
});

test('authored poster is reused and invalid or absent poster blocks output', async () => {
  const wave = await sample(), png = await readFile(new URL('../../../tests/fixtures/transparent-replacement.png', import.meta.url));
  for (const broken of [false, true]) {
    const p = createPresentation(); p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: wave } });
    p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'bytes', value: broken ? Uint8Array.of(1) : png } });
    p.addSlide({ elements: [{ type: 'audio', audioAssetId: 'sound', posterAssetId: 'poster', box: { x: 0, y: 0, width: 64, height: 64 } }] });
    if (broken) await assert.rejects(generatePptx(p, { audioMode: 'wave-pcm-v1' }), /poster could not be packaged/);
    else {
      const parts = readGeneratedZip((await generatePptx(p, { audioMode: 'wave-pcm-v1' })).bytes);
      assert.deepEqual(parts.get('ppt/media/poster.png'), png);
      assert.equal(parts.has('ppt/media/pptkit-audio-poster.png'), false);
    }
  }
});

test('PPTX rejects audio before resolving any assets instead of silently dropping it', async () => {
  const p = createPresentation();
  p.registerAsset({ id: 'audio', kind: 'audio', source: { type: 'url', value: 'https://example.invalid/audio.wav' } });
  p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'url', value: 'https://example.invalid/poster.png' } });
  p.addSlide({ elements: [{ type: 'audio', audioAssetId: 'audio', posterAssetId: 'poster', box: { x: 0, y: 0, width: 100, height: 50 } }] });
  let calls = 0;
  await assert.rejects(generatePptx(p, { resolveAsset: async () => { calls++; throw Error('unexpected'); } }), /requires audioMode wave-pcm-v1/);
  assert.equal(calls, 0);
});

test('hidden nested audio is rejected even though ordinary hidden objects are omitted', async () => {
  const p = createPresentation();
  p.registerAsset({ id: 'audio', kind: 'audio', source: { type: 'bytes', value: Uint8Array.of(1) } });
  p.addSlide({ hidden: true, elements: [{ type: 'group', box: { x: 0, y: 0, width: 100, height: 50 }, coordinateSize: { width: 100, height: 50 }, children: [{ type: 'audio', hidden: true, audioAssetId: 'audio', box: { x: 0, y: 0, width: 20, height: 20 } }] }] });
  await assert.rejects(generatePptx(p), /requires audioMode wave-pcm-v1/);
});
