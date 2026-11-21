import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { createPresentation } from '../packages/core/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { createServer } from 'node:http';
import { once } from 'node:events';

test('video source projection and static preview match Node and browser without media fetches', async t => {
  const bytes = await readFile('tests/fixtures/h264-mp4-powerpoint.pptx');
  const node = await importPptx(bytes, { projection: 'content', videoProjection: 'iso-bmff-v1' });
  let expected;
  try {
    const rendered = await renderPresentationToSvg(node.document);
    await assert.rejects(generatePptx(node.document), /requires videoMode iso-bmff-v1/);
    const exported = await generatePptx(node.document, { videoMode: 'iso-bmff-v1' });
    expected = { sources: node.report.videoSources, elements: node.document.slides[0].elements, rendered, exported: { bytes: [...exported.bytes], warnings: exported.warnings } };
  } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-parser', 'svg-renderer', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async data => {
    const session = await PPTKit.importPptx(Uint8Array.from(data), { projection: 'content', videoProjection: 'iso-bmff-v1' });
    try {
      const rendered = await PPTKit.renderPresentationToSvg(session.document);
      let error = ''; try { await PPTKit.generatePptx(session.document); } catch (e) { error = e.message; }
      const exported = await PPTKit.generatePptx(session.document, { videoMode: 'iso-bmff-v1' });
      return { sources: session.report.videoSources, elements: session.document.slides[0].elements, rendered, exported: { bytes: [...exported.bytes], warnings: exported.warnings }, error };
    } finally { session.dispose(); }
  }, [...bytes]);
  const { error, ...actualData } = actual;
  assert.match(error, /requires videoMode iso-bmff-v1/); assert.deepEqual(actualData, expected); assert.deepEqual(requests, []);
});

test('same-slide shared video re-export agrees across Node and browser', async t => {
  const bytes = await readFile('tests/fixtures/h264-mp4-pptkit-multi.pptx');
  const node = await importPptx(bytes, { projection: 'content', videoProjection: 'iso-bmff-v1' });
  let expected; try { const result = await generatePptx(node.document, { videoMode: 'iso-bmff-v1' }); expected = { bytes: [...result.bytes], warnings: result.warnings }; } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-parser', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async data => {
    const s = await PPTKit.importPptx(Uint8Array.from(data), { projection: 'content', videoProjection: 'iso-bmff-v1' });
    try { const result = await PPTKit.generatePptx(s.document, { videoMode: 'iso-bmff-v1' }); return { bytes: [...result.bytes], warnings: result.warnings }; } finally { s.dispose(); }
  }, [...bytes]);
  assert.deepEqual(actual, expected); assert.deepEqual(requests, []);
});

for (const [fixture, starts] of [['pptkit-mixed-av', ['click-sequence', 'click-sequence']], ['powerpoint-audio-start-auto', ['click-sequence', 'automatic']]]) test(`mixed media ${fixture} projection and re-export agree across Node and browser`, async t => {
  const bytes = await readFile(`tests/fixtures/${fixture}.pptx`);
  const importOptions = { projection: 'content', videoProjection: 'iso-bmff-v1', audioProjection: 'wave-pcm-v1' };
  const exportOptions = { videoMode: 'iso-bmff-v1', audioMode: 'wave-pcm-v1' };
  const node = await importPptx(bytes, importOptions);
  let expected;
  try {
    const result = await generatePptx(node.document, exportOptions);
    const plan = await node.prepareSave({ mode: 'no-op' });
    expected = { elements: node.document.slides[0].elements, audio: node.report.audioSources, video: node.report.videoSources, bytes: [...result.bytes], warnings: result.warnings, original: [...(await node.save(plan.plan)).bytes] };
  } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-parser', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async ({ data, importOptions, exportOptions }) => {
    const session = await PPTKit.importPptx(Uint8Array.from(data), importOptions);
    try {
      const result = await PPTKit.generatePptx(session.document, exportOptions);
      const plan = await session.prepareSave({ mode: 'no-op' });
      return { elements: session.document.slides[0].elements, audio: session.report.audioSources, video: session.report.videoSources, bytes: [...result.bytes], warnings: result.warnings, original: [...(await session.save(plan.plan)).bytes] };
    } finally { session.dispose(); }
  }, { data: [...bytes], importOptions, exportOptions });
  assert.deepEqual(actual, expected); assert.deepEqual(actual.original, [...bytes]); assert.deepEqual(actual.elements.filter(x => x.type === "audio").map(x => x.playbackStart), starts); assert.deepEqual(requests, []);
});

test('native trimmed audio import and fresh export agree across Node and browser', async t => {
  const bytes = await readFile('tests/fixtures/powerpoint-audio-trim.pptx');
  const importOptions = { projection: 'content', audioProjection: 'wave-pcm-v1' }, exportOptions = { audioMode: 'wave-pcm-v1' };
  const node = await importPptx(bytes, importOptions); let expected;
  try { const result = await generatePptx(node.document, exportOptions); const plan = await node.prepareSave({ mode: 'no-op' }); expected = { elements: node.document.slides[0].elements, sources: node.report.audioSources, bytes: [...result.bytes], warnings: result.warnings, original: [...(await node.save(plan.plan)).bytes] }; } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-parser', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async ({ data, importOptions, exportOptions }) => {
    const session = await PPTKit.importPptx(Uint8Array.from(data), importOptions);
    try { const result = await PPTKit.generatePptx(session.document, exportOptions), plan = await session.prepareSave({ mode: 'no-op' }); return { elements: session.document.slides[0].elements, sources: session.report.audioSources, bytes: [...result.bytes], warnings: result.warnings, original: [...(await session.save(plan.plan)).bytes] }; } finally { session.dispose(); }
  }, { data: [...bytes], importOptions, exportOptions });
  assert.deepEqual(actual, expected); assert.deepEqual(actual.original, [...bytes]); assert.deepEqual(requests, []);
  assert.deepEqual(actual.elements.filter(x => x.type === 'audio').map(x => [x.playbackVolume, x.loop, x.rewindAfterPlayback]), [['high', false, false], ['high', true, false], ['high', false, true], ['high', true, true]]);
  assert.deepEqual(actual.elements.filter(x => x.type === 'audio').map(x => x.playbackStart), ['click-sequence', 'object-click', 'click-sequence', 'object-click']);
});

test('native automatic audio main sequence agrees across Node and browser', async t => {
  const bytes = await readFile('tests/fixtures/powerpoint-two-audio-auto.pptx'), options = { projection: 'content', audioProjection: 'wave-pcm-v1' };
  const node = await importPptx(bytes, options); let expected;
  try { const output = await generatePptx(node.document, { audioMode: 'wave-pcm-v1' }); expected = { elements: node.document.slides[0].elements, bytes: [...output.bytes], warnings: output.warnings }; } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-parser', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async ({ bytes, options }) => { const s = await PPTKit.importPptx(Uint8Array.from(bytes), options); try { const output = await PPTKit.generatePptx(s.document, { audioMode: 'wave-pcm-v1' }); return { elements: s.document.slides[0].elements, bytes: [...output.bytes], warnings: output.warnings }; } finally { s.dispose(); } }, { bytes: [...bytes], options });
  assert.deepEqual(actual, expected); assert.deepEqual(actual.elements.filter(x => x.type === 'audio').map(x => x.playbackStart), ['automatic', 'automatic']); assert.deepEqual(requests, []);
});

test('authored object-click audio export agrees across Node and browser', async t => {
  const parts = readGeneratedZip(await readFile('tests/fixtures/pptkit-audio-start-object-click.pptx'));
  const wave = parts.get('ppt/media/sound.wav');
  const make = create => { const p = create(); p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: wave } }); p.addSlide({ elements: [{ type: 'audio', audioAssetId: 'sound', playbackStart: 'object-click', playbackVolume: 'mute', loop: true, rewindAfterPlayback: true, trimStartMs: 100, trimEndMs: 50, box: { x: 0, y: 0, width: 40, height: 40 } }, { type: 'audio', audioAssetId: 'sound', playbackStart: 'automatic', playbackVolume: 'low', loop: false, rewindAfterPlayback: true, trimEndMs: 200, box: { x: 50, y: 0, width: 40, height: 40 } }] }); return p; };
  const expected = await generatePptx(make(createPresentation), { audioMode: 'wave-pcm-v1' });
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async wave => {
    const p = PPTKit.createPresentation(); p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: Uint8Array.from(wave) } });
    p.addSlide({ elements: [{ type: 'audio', audioAssetId: 'sound', playbackStart: 'object-click', playbackVolume: 'mute', loop: true, rewindAfterPlayback: true, trimStartMs: 100, trimEndMs: 50, box: { x: 0, y: 0, width: 40, height: 40 } }, { type: 'audio', audioAssetId: 'sound', playbackStart: 'automatic', playbackVolume: 'low', loop: false, rewindAfterPlayback: true, trimEndMs: 200, box: { x: 50, y: 0, width: 40, height: 40 } }] });
    const result = await PPTKit.generatePptx(p, { audioMode: 'wave-pcm-v1' }); return { bytes: [...result.bytes], warnings: result.warnings };
  }, [...wave]);
  assert.deepEqual(actual.bytes, [...expected.bytes]); assert.deepEqual(actual.warnings, expected.warnings); assert.deepEqual(requests, []);
});

test('external mixed media resolver agrees across Node and browser without implicit network', async t => {
  const parts = readGeneratedZip(await readFile('tests/fixtures/pptkit-mixed-av.pptx'));
  const media = { movie: [...parts.get('ppt/media/movie.mp4')], sound: [...parts.get('ppt/media/sound.wav')] };
  const make = create => {
    const p = create();
    p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'url', value: 'https://example.invalid/movie.mp4' } });
    p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'url', value: 'https://example.invalid/sound.wav' } });
    p.addSlide({ elements: [{ type: 'video', videoAssetId: 'movie', box: { x: 0, y: 0, width: 40, height: 40 } }, { type: 'audio', audioAssetId: 'sound', box: { x: 50, y: 0, width: 40, height: 40 } }] });
    return p;
  };
  const options = { videoMode: 'iso-bmff-v1', audioMode: 'wave-pcm-v1', maxAssetBytes: 100000 };
  const expected = await generatePptx(make(createPresentation), { ...options, resolveMedia: asset => ({ bytes: Uint8Array.from(media[asset.id]) }) });
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async ({ media, options }) => {
    const p = PPTKit.createPresentation();
    p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'url', value: 'https://example.invalid/movie.mp4' } });
    p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'url', value: 'https://example.invalid/sound.wav' } });
    p.addSlide({ elements: [{ type: 'video', videoAssetId: 'movie', box: { x: 0, y: 0, width: 40, height: 40 } }, { type: 'audio', audioAssetId: 'sound', box: { x: 50, y: 0, width: 40, height: 40 } }] });
    const calls = []; const result = await PPTKit.generatePptx(p, { ...options, resolveMedia: (asset, context) => { calls.push([asset.id, context.maxBytes]); return { bytes: Uint8Array.from(media[asset.id]) }; } });
    return { bytes: [...result.bytes], warnings: result.warnings, calls };
  }, { media, options });
  assert.deepEqual(actual.bytes, [...expected.bytes]); assert.deepEqual(actual.warnings, expected.warnings);
  assert.deepEqual(actual.calls, [['movie', 100000], ['sound', 100000]]); assert.deepEqual(requests, []);
});

test('HTTP media adapter streams in Chromium and refuses opaque redirects', async t => {
  const parts = readGeneratedZip(await readFile('tests/fixtures/pcm-wave-powerpoint.pptx'));
  const wave = parts.get('ppt/media/media1.wav'), hits = [];
  const server = createServer((req, res) => {
    hits.push(req.url);
    if (req.url === '/redirect') { res.writeHead(302, { location: '/sound' }); res.end(); return; }
    if (req.url === '/sound') { res.writeHead(200, { 'content-type': 'audio/wav' }); res.write(wave.subarray(0, 100)); res.end(wave.subarray(100)); return; }
    res.writeHead(200, { 'content-type': 'text/html' }); res.end('<!doctype html>');
  });
  server.listen(0, '127.0.0.1'); await once(server, 'listening');
  t.after(() => { server.closeAllConnections(); server.close(); });
  const base = `http://127.0.0.1:${server.address().port}`;
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.goto(base);
  for (const name of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const result = await page.evaluate(async ({ base, limit }) => {
    const make = path => { const p = PPTKit.createPresentation(); p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'url', value: base + path } }); p.addSlide({ elements: [{ type: 'audio', audioAssetId: 'sound', box: { x: 0, y: 0, width: 40, height: 40 } }] }); return p; };
    const calls = [], resolveMedia = PPTKit.createHttpMediaResolver({ maxRedirects: 1, authorizeUrl: url => { calls.push(url); return true; } });
    const options = { audioMode: 'wave-pcm-v1', maxAssetBytes: limit, resolveMedia };
    const ok = await PPTKit.generatePptx(make('/sound'), options);
    let redirectError = '', sizeError = '';
    try { await PPTKit.generatePptx(make('/redirect'), options); } catch (e) { redirectError = e.message; }
    try { await PPTKit.generatePptx(make('/sound'), { ...options, maxAssetBytes: limit - 1 }); } catch (e) { sizeError = e.message; }
    return { bytes: [...ok.bytes], calls, redirectError, sizeError };
  }, { base, limit: wave.length });
  assert.deepEqual(readGeneratedZip(Uint8Array.from(result.bytes)).get('ppt/media/sound.wav'), wave);
  assert.deepEqual(result.calls, [base + '/sound', base + '/redirect', base + '/sound']);
  assert.match(result.redirectError, /not an accessible 200/); assert.match(result.sizeError, /maxBytes/);
  assert.equal(hits.filter(path => path === '/sound').length, 2);
});
