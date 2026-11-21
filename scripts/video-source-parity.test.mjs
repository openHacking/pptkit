import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { chromium } from '@playwright/test';
import { importPptx, inspectPptxCaptionTracks } from '../packages/pptx-parser/dist/index.js';

test('video source inventory and no-op bytes agree in Node and browser', async t => {
  const bytes = await readFile('tests/fixtures/h264-mp4-powerpoint.pptx');
  const { unzipSync, zipSync, strFromU8, strToU8 } = createRequire(new URL('../packages/pptx-parser/package.json', import.meta.url))('fflate');
  const parts = unzipSync(bytes), rels = 'ppt/slides/_rels/slide1.xml.rels';
  const original = 'Target="../media/media1.mp4"/><Relationship Id="rId1"';
  const xml = strFromU8(parts[rels]); assert.ok(xml.includes(original));
  parts[rels] = strToU8(xml.replace(original, 'Target="https://example.invalid/video.mp4" TargetMode="External"/><Relationship Id="rId1"'));
  const samples = [bytes, zipSync(parts)];
  const options = { projection: 'content', videoSourceProfile: 'iso-bmff-v1' };
  const expected = [];
  for (const sample of samples) {
    const node = await importPptx(sample, options);
    try { expected.push(node.report.videoSources); } finally { node.dispose(); }
  }
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const core = await readFile('packages/core/dist/browser/index.js');
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': `data:text/javascript;base64,${core.toString('base64')}` } })}</script>`);
  await page.addScriptTag({ content: await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8') });
  const bundle = await readFile('packages/pptx-parser/dist/browser/index.js');
  const actual = await page.evaluate(async ({ samples, url }) => {
    const esm = await import(url), results = [];
    for (const data of samples) for (const api of [esm, globalThis.PPTKit]) {
      const session = await api.importPptx(Uint8Array.from(data), { projection: 'content', videoSourceProfile: 'iso-bmff-v1' });
      try {
        const plan = await session.prepareSave({ mode: 'no-op' }), saved = await session.save(plan.plan);
        results.push({ sources: session.report.videoSources, same: saved.bytes.length === data.length && saved.bytes.every((b, i) => b === data[i]) });
      } finally { session.dispose(); }
    }
    return results;
  }, { samples: samples.map(sample => [...sample]), url: `data:text/javascript;base64,${bundle.toString('base64')}` });
  assert.deepEqual(actual, expected.flatMap(sources => [{ sources, same: true }, { sources, same: true }]));
  assert.deepEqual(requests, []);
});

test('SRT-origin caption source no-op and unsupported findings agree in Node and browser', async t => {
  const bytes = await readFile('tests/fixtures/powerpoint-caption-srt.pptx');
  const options = { projection: 'content', videoProjection: 'iso-bmff-v1' };
  const node = await importPptx(bytes, options); let expected;
  try {
    const plan = await node.prepareSave({ mode: 'no-op' });
    expected = { videoSources: node.report.videoSources, findings: node.report.findings.filter(x => x.feature === 'source-part-content' && x.operation === 'import' && x.source.part.endsWith('.vtt')), bytes: [...(await node.save(plan.plan)).bytes] };
  } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async ({ data, options }) => {
    const session = await PPTKit.importPptx(Uint8Array.from(data), options);
    try { const plan = await session.prepareSave({ mode: 'no-op' }); return { videoSources: session.report.videoSources, findings: session.report.findings.filter(x => x.feature === 'source-part-content' && x.operation === 'import' && x.source.part.endsWith('.vtt')), bytes: [...(await session.save(plan.plan)).bytes] }; }
    finally { session.dispose(); }
  }, { data: [...bytes], options });
  assert.deepEqual(actual, expected); assert.equal(actual.findings.length, 3); assert.deepEqual(actual.bytes, [...bytes]); assert.deepEqual(requests, []);
});

test('caption source inventory matches Node, browser ESM and global without external requests', async t => {
  const bytes = await readFile('tests/fixtures/powerpoint-caption-srt.pptx');
  const { unzipSync, zipSync, strFromU8, strToU8 } = createRequire(new URL('../packages/pptx-parser/package.json', import.meta.url))('fflate');
  const parts = unzipSync(bytes), rels = 'ppt/slides/_rels/slide1.xml.rels';
  parts[rels] = strToU8(strFromU8(parts[rels]).replace('Target="../media/track1.vtt"', 'Target="https://example.invalid/track.vtt" TargetMode="External"'));
  const inactive = unzipSync(bytes), slide = 'ppt/slides/slide1.xml';
  inactive[slide] = strToU8(strFromU8(inactive[slide]).replace('<p:pic>', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:unsupported"><mc:Choice Requires="f"><p:pic>').replace('</p:pic>', '</p:pic></mc:Choice><mc:Fallback/></mc:AlternateContent>'));
  const wrapped = unzipSync(bytes), xml = strFromU8(wrapped[slide]), track = xml.match(/<p173:track id="\{363C0747[^>]*\/>/u)?.[0]; assert.ok(track);
  wrapped[slide] = strToU8(xml.replace(track, `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:unsupported"><mc:Choice Requires="f">${track}</mc:Choice><mc:Fallback>${track}</mc:Fallback></mc:AlternateContent>`));
  const samples = [bytes, zipSync(parts), zipSync(inactive), zipSync(wrapped)], expected = await Promise.all(samples.map(sample => inspectPptxCaptionTracks(sample)));
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const core = await readFile('packages/core/dist/browser/index.js');
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': `data:text/javascript;base64,${core.toString('base64')}` } })}</script>`);
  await page.addScriptTag({ content: await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8') });
  const bundle = await readFile('packages/pptx-parser/dist/browser/index.js');
  const actual = await page.evaluate(async ({ samples, url }) => {
    const esm = await import(url), out = [];
    for (const bytes of samples) for (const api of [esm, globalThis.PPTKit]) out.push(await api.inspectPptxCaptionTracks(Uint8Array.from(bytes)));
    return out;
  }, { samples: samples.map(sample => [...sample]), url: `data:text/javascript;base64,${bundle.toString('base64')}` });
  assert.deepEqual(actual, expected.flatMap(result => [result, result])); assert.equal(actual[2].tracks[0].status, 'external'); assert.equal(actual[4].tracks[0].pictureView, 'raw-only'); assert.equal(actual[6].tracks.length, 4); assert.deepEqual(actual[6].tracks.slice(0, 2).map(track => track.trackIdStatus), ['repeated', 'repeated']); assert.deepEqual(requests, []);
});
