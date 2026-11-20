import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

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

test('mixed audio and video projection and re-export agree across Node and browser', async t => {
  const bytes = await readFile('tests/fixtures/pptkit-mixed-av.pptx');
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
  assert.deepEqual(actual, expected); assert.deepEqual(actual.original, [...bytes]); assert.deepEqual(requests, []);
});
