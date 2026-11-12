import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { imageFixture, encodeImages, picture } from '../packages/pptx-parser/test/image-fixture.mjs';
import { relationships } from '../packages/pptx-parser/test/opc-fixture.mjs';

test('picture source inventory matches browser ESM/global and independent raw DOM', async t => {
  const extension = '<a:extLst><a:ext uri="svg"><v:svgBlip xmlns:v="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="svg"/></a:ext></a:extLst>';
  const fixture = imageFixture(picture({ id: 2, link: 'outside', effects: extension }) + picture({ id: 3 }));
  fixture.parts['media/image.svg'] = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>'); fixture.mediaTypes['media/image.svg'] = 'image/svg+xml';
  fixture.parts['deck/_rels/z.xml.rels'] = relationships([fixture.rel('layout', 'slideLayout', 'layout.xml'), fixture.rel('image', 'image', '../media/raster.bin'), fixture.rel('svg', 'image', '../media/image.svg'), { ...fixture.rel('outside', 'image', 'https://example.invalid/image.png'), TargetMode: 'External' }]);
  const bytes = encodeImages(fixture);
  const node = await importPptx(bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  t.after(() => node.dispose());
  const esm = await readFile('packages/pptx-parser/dist/browser/index.js'), global = await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8');
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: global });
  const results = await page.evaluate(async ({ url, data, xml, rel }) => {
    const module = await import(url), doc = new DOMParser().parseFromString(xml, 'application/xml'), elements = Array.from(doc.getElementsByTagName('*'));
    const results = [];
    for (const api of [module, PPTKit]) {
      const session = await api.importPptx(new Uint8Array(data), { projection: 'content', pictureSourceProfile: 'embedded-v1' });
      const sources = session.report.pictureSources;
      results.push({ sources, raw: sources.map(source => { const node = elements[source.sourceElementIndex - 1]; return { local: node.localName, id: node.getAttributeNS(rel, 'embed') }; }) });
      session.dispose();
    }
    return results;
  }, { url: `data:text/javascript;base64,${esm.toString('base64')}`, data: Array.from(bytes), xml: fixture.parts['deck/z.xml'], rel: fixture.R });
  for (const result of results) { assert.deepEqual(result.sources, node.report.pictureSources); assert.deepEqual(result.raw, [{ local: 'blip', id: 'image' }, { local: 'blip', id: 'image' }]); }
  assert.equal(node.report.pictureSources[0].hasSvgExtension, true); assert.equal(node.report.pictureSources[0].hasExternalLink, true);
  assert.deepEqual(requests, []);
});

test('picture inventory composes with starting-number preservation in Node, ESM and global', async t => {
  const input = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))), options = { projection: 'content', pictureSourceProfile: 'embedded-v1', firstSlideNumberSaveProfile: 'attribute-v2' };
  const node = await importPptx(input, options); let expected;
  try { node.document.setFirstSlideNumber(17); const plan = await node.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); expected = { sources: node.report.pictureSources, bytes: [...(await node.save(plan.plan)).bytes] }; } finally { node.dispose(); }
  const esm = await readFile('packages/pptx-parser/dist/browser/index.js'), global = await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8');
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: global });
  const result = await page.evaluate(async ({ url, data, options }) => { const outputs = []; for (const api of [await import(url), PPTKit]) { const session = await api.importPptx(Uint8Array.from(data), options); try { session.document.setFirstSlideNumber(17); const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw Error(JSON.stringify(plan.report)); outputs.push({ sources: session.report.pictureSources, bytes: [...(await session.save(plan.plan)).bytes] }); } finally { session.dispose(); } } return outputs; }, { url: `data:text/javascript;base64,${esm.toString('base64')}`, data: [...input], options });
  for (const output of result) assert.deepEqual(output, expected); assert.deepEqual(requests, []);
});
