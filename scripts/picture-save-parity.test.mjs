import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { deflateSync } from 'node:zlib';
import { importPptx, inspectOpcPackage } from '../packages/pptx-parser/dist/index.js';
import { imageFixture, encodeImages, picture, validGif as gif, gif as incompleteGif, jpegHeader, png } from '../packages/pptx-parser/test/image-fixture.mjs';
import { rasterMetadata } from '../packages/pptx-parser/dist/import/image-metadata.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { indexedRows, filterRows } from '../tests/support/png-indexed-fixture.mjs';

test('public picture save matches Node, browser ESM/global and independent source XML', async t => {
  const source = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))), options = { projection: 'content', pictureEditSaveProfile: 'replace-v1' };
  const session = await importPptx(source, options); t.after(() => session.dispose());
  const slide = session.document.slides[0], element = slide.elements[0];
  session.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: gif }, mimeType: 'image/gif', width: 1, height: 1 });
  slide.replaceElement(element.id, { ...element, assetId: 'replacement' });
  const prepared = await session.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'ready'); const expected = await session.save(prepared.plan);
  const esm = await readFile('packages/pptx-parser/dist/browser/index.js'), global = await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8');
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: global });
  const results = await page.evaluate(async ({ url, bytes, image, options }) => {
    const module = await import(url), result = [];
    for (const api of [module, PPTKit]) {
      const s = await api.importPptx(new Uint8Array(bytes), options), slide = s.document.slides[0], element = slide.elements[0];
      s.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: new Uint8Array(image) }, mimeType: 'image/gif', width: 1, height: 1 });
      slide.replaceElement(element.id, { ...element, assetId: 'replacement' });
      const prepared = await s.prepareSave({ mode: 'edit' }); if (prepared.status !== 'ready') throw Error(JSON.stringify(prepared.report));
      const output = await s.save(prepared.plan); result.push({ bytes: Array.from(output.bytes), report: output.report }); s.dispose();
    }
    return result;
  }, { url: `data:text/javascript;base64,${esm.toString('base64')}`, bytes: Array.from(source), image: Array.from(gif), options });
  for (const result of results) { assert.deepEqual(new Uint8Array(result.bytes), expected.bytes); assert.deepEqual(result.report, expected.report); }
  const parts = readGeneratedZip(expected.bytes), slideXml = parts.get('deck/z.xml').toString(), relXml = parts.get('deck/_rels/z.xml.rels').toString();
  const inspection = await inspectOpcPackage(expected.bytes), added = inspection.relationships.find(item => item.source === '/deck/z.xml' && item.id.startsWith('pptkitImage'));
  assert.ok(added); assert.ok(slideXml.includes(`r:embed="${added.id}"`)); assert.ok(slideXml.includes('r:embed="image"')); assert.ok(relXml.includes(`Id="${added.id}"`));
  assert.deepEqual(requests, []);
});

test('incomplete GIF rejection and corrected retry agree across Node and browser entries', async t => {
  const source = encodeImages(), options = { projection: 'content', pictureEditSaveProfile: 'replace-v1' };
  const run = async (api, input, bad, good, options) => {
    const s = await api.importPptx(new Uint8Array(input), options), slide = s.document.slides[0], element = slide.elements[0];
    try {
      s.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: new Uint8Array(bad) }, mimeType: 'image/gif', width: 1, height: 1 });
      slide.replaceElement(element.id, { ...element, assetId: 'replacement' });
      const rejected = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(good) }, mimeType: 'image/gif', width: 1, height: 1 });
      const ready = await s.prepareSave({ mode: 'edit' });
      if (ready.status !== 'ready') throw Error(JSON.stringify(ready.report));
      return { code: rejected.report.diagnostics[0].code, bytes: Array.from((await s.save(ready.plan)).bytes) };
    } finally { s.dispose(); }
  };
  const expected = await run({ importPptx }, source, incompleteGif, gif, options);
  const esm = await readFile('packages/pptx-parser/dist/browser/index.js'), global = await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8');
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: global });
  const actual = await page.evaluate(async ({ url, source, bad, good, options, runSource }) => {
    const module = await import(url), run = (0, eval)(`(${runSource})`);
    return Promise.all([module, PPTKit].map(api => run(api, source, bad, good, options)));
  }, { url: `data:text/javascript;base64,${esm.toString('base64')}`, source: Array.from(source), bad: Array.from(incompleteGif), good: Array.from(gif), options, runSource: run.toString() });
  assert.equal(expected.code, 'save-picture-candidate-unsupported');
  for (const result of actual) assert.deepEqual(result, expected);
  assert.deepEqual(requests, []);
});

test('incomplete JPEG rejection and corrected retry agree across Node and browser entries', async t => {
  const source = encodeImages(), options = { projection: 'content', pictureEditSaveProfile: 'replace-v1' };
  const good = await readFile('tests/fixtures/jpeg-sequential-multiscan.jpg'), { width, height } = await rasterMetadata(good, async () => {});
  const badTable = Buffer.from(good); badTable[badTable.indexOf(Buffer.from([0xff, 0xdb])) + 5] = 0;
  const badCoverage = Buffer.concat([good.subarray(0, good.lastIndexOf(Buffer.from([0xff, 0xda]))), Buffer.from([0xff, 0xd9])]);
  const progressive = await readFile('tests/fixtures/jpeg-progressive-components.jpg');
  const twelve = await readFile('tests/fixtures/jpeg-12bit-sequential.jpg');
  const badPrecision = Buffer.from(good); badPrecision[badPrecision.indexOf(Buffer.from([0xff, 0xc0])) + 4] = 12;
  const unordered = Buffer.from(progressive), firstScan = unordered.indexOf(Buffer.from([0xff, 0xda]));
  unordered[firstScan + 7 + 2 * unordered[firstScan + 4]] = 0x21;
  const run = async (api, input, bad, badTable, badCoverage, badPrecision, twelve, unordered, progressive, good, options, width, height) => {
    const s = await api.importPptx(new Uint8Array(input), options), slide = s.document.slides[0], element = slide.elements[0];
    try {
      s.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: new Uint8Array(bad) }, mimeType: 'image/jpeg', width: 8, height: 4 });
      slide.replaceElement(element.id, { ...element, assetId: 'replacement' });
      const rejected = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(badTable) }, mimeType: 'image/jpeg', width, height });
      const rejectedTable = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(badCoverage) }, mimeType: 'image/jpeg', width, height });
      const rejectedCoverage = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(badPrecision) }, mimeType: 'image/jpeg', width, height });
      const rejectedPrecision = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(twelve) }, mimeType: 'image/jpeg', width, height });
      const twelveReady = await s.prepareSave({ mode: 'edit' });
      if (twelveReady.status === 'ready') twelveReady.plan.dispose();
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(unordered) }, mimeType: 'image/jpeg', width: 8, height: 8 });
      const rejectedOrder = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(progressive) }, mimeType: 'image/jpeg', width: 8, height: 8 });
      const progressiveReady = await s.prepareSave({ mode: 'edit' });
      if (progressiveReady.status === 'ready') progressiveReady.plan.dispose();
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(good) }, mimeType: 'image/jpeg', width, height });
      const ready = await s.prepareSave({ mode: 'edit' }); if (ready.status !== 'ready') throw Error(JSON.stringify(ready.report));
      return { code: rejected.report.diagnostics[0].code, tableCode: rejectedTable.report.diagnostics[0].code, coverageCode: rejectedCoverage.report.diagnostics[0].code, precisionCode: rejectedPrecision.report.diagnostics[0].code, twelveStatus: twelveReady.status, orderCode: rejectedOrder.report.diagnostics[0].code, progressiveStatus: progressiveReady.status, bytes: Array.from((await s.save(ready.plan)).bytes) };
    } finally { s.dispose(); }
  };
  const expected = await run({ importPptx }, source, jpegHeader, badTable, badCoverage, badPrecision, twelve, unordered, progressive, good, options, width, height);
  const esm = await readFile('packages/pptx-parser/dist/browser/index.js'), global = await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8');
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: global });
  const actual = await page.evaluate(async ({ url, source, bad, badTable, badCoverage, badPrecision, twelve, unordered, progressive, good, options, width, height, runSource }) => {
    const module = await import(url), run = (0, eval)(`(${runSource})`);
    return Promise.all([module, PPTKit].map(api => run(api, source, bad, badTable, badCoverage, badPrecision, twelve, unordered, progressive, good, options, width, height)));
  }, { url: `data:text/javascript;base64,${esm.toString('base64')}`, source: Array.from(source), bad: Array.from(jpegHeader), badTable: Array.from(badTable), badCoverage: Array.from(badCoverage), badPrecision: Array.from(badPrecision), twelve: Array.from(twelve), unordered: Array.from(unordered), progressive: Array.from(progressive), good: Array.from(good), options, width, height, runSource: run.toString() });
  assert.equal(expected.code, 'save-picture-candidate-unsupported'); assert.equal(expected.tableCode, 'save-picture-candidate-unsupported');
  assert.equal(expected.coverageCode, 'save-picture-candidate-unsupported');
  assert.equal(expected.precisionCode, 'save-picture-candidate-unsupported'); assert.equal(expected.twelveStatus, 'ready');
  assert.equal(expected.orderCode, 'save-picture-candidate-unsupported'); assert.equal(expected.progressiveStatus, 'ready');
  for (const result of actual) assert.deepEqual(result, expected);
  assert.deepEqual(requests, []);
});

test('PNG container failure and corrected retry agree across Node and browser entries', async t => {
  const source = encodeImages(), options = { projection: 'content', pictureEditSaveProfile: 'replace-v1' }, truncated = png.subarray(0, 33);
  const badStream = Buffer.from(png), chunkStart = 33, length = badStream.readUInt32BE(chunkStart), chunkEnd = chunkStart + 8 + length;
  badStream[chunkEnd - 1] ^= 1; let crc = 0xffffffff;
  for (const byte of badStream.subarray(chunkStart + 4, chunkEnd)) { crc ^= byte; for (let i = 0; i < 8; i++) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1; }
  badStream.writeUInt32BE((crc ^ 0xffffffff) >>> 0, chunkEnd);
  const chunk = (type, body) => { const result = Buffer.alloc(12 + body.length); result.writeUInt32BE(body.length); result.write(type, 4); body.copy(result, 8); let checksum = 0xffffffff; for (const byte of result.subarray(4, 8 + body.length)) { checksum ^= byte; for (let i = 0; i < 8; i++) checksum = checksum & 1 ? 0xedb88320 ^ (checksum >>> 1) : checksum >>> 1; } result.writeUInt32BE((checksum ^ 0xffffffff) >>> 0, 8 + body.length); return result; };
  const ihdr = Buffer.from(png.subarray(16, 29)); ihdr.writeUInt32BE(1, 0); ihdr.writeUInt32BE(1, 4); ihdr[8] = 8; ihdr[9] = 3; ihdr[12] = 0;
  const badPalette = Buffer.concat([png.subarray(0, 8), chunk('IHDR', ihdr), chunk('PLTE', Buffer.from([0, 0, 0])), chunk('IDAT', deflateSync(Buffer.from([0, 1]))), chunk('IEND', Buffer.alloc(0))]);
  const badTransparency = Buffer.concat([png.subarray(0, 8), chunk('IHDR', ihdr), chunk('PLTE', Buffer.from([0, 0, 0])), chunk('tRNS', Buffer.from([0, 255])), chunk('IDAT', deflateSync(Buffer.from([0, 0]))), chunk('IEND', Buffer.alloc(0))]);
  const validHeader = Buffer.from(ihdr); validHeader.writeUInt32BE(11, 0); validHeader.writeUInt32BE(5, 4); validHeader[8] = 2;
  const validIndexed = Buffer.concat([png.subarray(0, 8), chunk('IHDR', validHeader), chunk('PLTE', Buffer.from([0, 0, 0, 255, 255, 255])), chunk('IDAT', deflateSync(filterRows(indexedRows(11, 5, 2)))), chunk('IEND', Buffer.alloc(0))]);
  const run = async api => {
    const s = await api.importPptx(source, options), slide = s.document.slides[0], element = slide.elements[0];
    try {
      s.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: truncated }, mimeType: 'image/png', width: 8, height: 4 });
      slide.replaceElement(element.id, { ...element, assetId: 'replacement' });
      const rejected = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: badStream }, mimeType: 'image/png', width: 8, height: 4 });
      const badCompressed = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: badPalette }, mimeType: 'image/png', width: 1, height: 1 });
      const badIndexed = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: badTransparency }, mimeType: 'image/png', width: 1, height: 1 });
      const badAlpha = await s.prepareSave({ mode: 'edit' });
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: validIndexed }, mimeType: 'image/png', width: 11, height: 5 });
      const indexed = await s.prepareSave({ mode: 'edit' }); assert.equal(indexed.status, 'ready'); const indexedBytes = Array.from((await s.save(indexed.plan)).bytes); indexed.plan.dispose();
      s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: png }, mimeType: 'image/png', width: 8, height: 4 });
      const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready'); return { code: rejected.report.diagnostics[0].code, compressedCode: badCompressed.report.diagnostics[0].code, paletteCode: badIndexed.report.diagnostics[0].code, alphaCode: badAlpha.report.diagnostics[0].code, indexedBytes, bytes: Array.from((await s.save(ready.plan)).bytes) };
    } finally { s.dispose(); }
  };
  const expected = await run({ importPptx }), esm = await readFile('packages/pptx-parser/dist/browser/index.js'), global = await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8');
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: global });
  const actual = await page.evaluate(async ({ url, source, truncated, badStream, badPalette, badTransparency, validIndexed, png, options }) => {
    const module = await import(url), results = [];
    for (const api of [module, PPTKit]) {
      const s = await api.importPptx(new Uint8Array(source), options), slide = s.document.slides[0], element = slide.elements[0];
      try {
        s.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: new Uint8Array(truncated) }, mimeType: 'image/png', width: 8, height: 4 });
        slide.replaceElement(element.id, { ...element, assetId: 'replacement' });
        const rejected = await s.prepareSave({ mode: 'edit' });
        s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(badStream) }, mimeType: 'image/png', width: 8, height: 4 });
        const badCompressed = await s.prepareSave({ mode: 'edit' });
        s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(badPalette) }, mimeType: 'image/png', width: 1, height: 1 });
        const badIndexed = await s.prepareSave({ mode: 'edit' });
        s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(badTransparency) }, mimeType: 'image/png', width: 1, height: 1 });
        const badAlpha = await s.prepareSave({ mode: 'edit' });
        s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(validIndexed) }, mimeType: 'image/png', width: 11, height: 5 });
        const indexed = await s.prepareSave({ mode: 'edit' }); if (indexed.status !== 'ready') throw Error(JSON.stringify(indexed.report)); const indexedBytes = Array.from((await s.save(indexed.plan)).bytes); indexed.plan.dispose();
        s.document.replaceAsset('replacement', { kind: 'image', source: { type: 'bytes', value: new Uint8Array(png) }, mimeType: 'image/png', width: 8, height: 4 });
        const ready = await s.prepareSave({ mode: 'edit' }); if (ready.status !== 'ready') throw Error(JSON.stringify(ready.report));
        results.push({ code: rejected.report.diagnostics[0].code, compressedCode: badCompressed.report.diagnostics[0].code, paletteCode: badIndexed.report.diagnostics[0].code, alphaCode: badAlpha.report.diagnostics[0].code, indexedBytes, bytes: Array.from((await s.save(ready.plan)).bytes) });
      } finally { s.dispose(); }
    }
    return results;
  }, { url: `data:text/javascript;base64,${esm.toString('base64')}`, source: Array.from(source), truncated: Array.from(truncated), badStream: Array.from(badStream), badPalette: Array.from(badPalette), badTransparency: Array.from(badTransparency), validIndexed: Array.from(validIndexed), png: Array.from(png), options });
  assert.equal(expected.code, 'save-picture-candidate-unsupported'); assert.equal(expected.compressedCode, 'save-picture-candidate-unsupported'); assert.equal(expected.paletteCode, 'save-picture-candidate-unsupported'); assert.equal(expected.alphaCode, 'save-picture-candidate-unsupported'); for (const result of actual) assert.deepEqual(result, expected); assert.deepEqual(requests, []);
});

test('combined picture replacement and starting number agree in Node, ESM and global', async t => {
  const source = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))), options = { projection: 'content', pictureEditSaveProfile: 'replace-v1', firstSlideNumberSaveProfile: 'attribute-v2' };
  const run = async (api, input, image, options) => { const s = await api.importPptx(new Uint8Array(input), options); try { const slide = s.document.slides[0], element = slide.elements[0]; s.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: new Uint8Array(image) }, mimeType: 'image/gif', width: 1, height: 1 }); slide.replaceElement(element.id, { ...element, assetId: 'replacement' }); s.document.setFirstSlideNumber(17); const p = await s.prepareSave({ mode: 'edit' }); if (p.status !== 'ready') throw Error(JSON.stringify(p.report)); const result = await s.save(p.plan); return { bytes: [...result.bytes], report: result.report }; } finally { s.dispose(); } };
  const expected = await run({ importPptx }, source, gif, options), esm = await readFile('packages/pptx-parser/dist/browser/index.js'), global = await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8');
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: global });
  const actual = await page.evaluate(async ({ url, source, image, options, runSource }) => { const run = (0, eval)(`(${runSource})`), module = await import(url); return Promise.all([module, PPTKit].map(api => run(api, source, image, options))); }, { url: `data:text/javascript;base64,${esm.toString('base64')}`, source: [...source], image: [...gif], options, runSource: run.toString() });
  for (const result of actual) assert.deepEqual(result, expected); assert.deepEqual(requests, []);
});

test('number-only edit with picture profile agrees in Node, ESM and global', async t => {
  const source = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))), options = { projection: 'content', pictureEditSaveProfile: 'replace-v1', firstSlideNumberSaveProfile: 'attribute-v2' };
  const run = async (api, input, options) => {
    let bytes = new Uint8Array(input); const results = [];
    for (const value of [17, 0, undefined]) {
      const s = await api.importPptx(bytes, options);
      try {
        s.document.setFirstSlideNumber(value); const p = await s.prepareSave({ mode: 'edit' });
        if (p.status !== 'ready') throw Error(JSON.stringify(p.report));
        const saved = await s.save(p.plan); bytes = saved.bytes; results.push({ bytes: [...bytes], report: saved.report });
      } finally { s.dispose(); }
    }
    return results;
  };
  const expected = await run({ importPptx }, source, options), esm = await readFile('packages/pptx-parser/dist/browser/index.js'), global = await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8');
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: global });
  const actual = await page.evaluate(async ({ url, source, options, runSource }) => { const run = (0, eval)(`(${runSource})`), module = await import(url); return Promise.all([module, PPTKit].map(api => run(api, source, options))); }, { url: `data:text/javascript;base64,${esm.toString('base64')}`, source: [...source], options, runSource: run.toString() });
  for (const result of actual) assert.deepEqual(result, expected); assert.deepEqual(requests, []);
});
