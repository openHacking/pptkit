import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { chromium } from '@playwright/test';
import { importSimpleSvg, importSvgWithFallback, createSvgRasterFallback } from '../packages/svg-parser/dist/index.js';
import { createPresentation } from '../packages/core/dist/index.js';
import { validatePresentation } from '../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { getSvgSlideSelectionGeometry, hitTestSvgSelectionGeometry } from '../packages/svg-renderer/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
const { unzipSync, zipSync, strFromU8, strToU8 } = createRequire(new URL('../packages/pptx-parser/package.json', import.meta.url))('fflate');

test('explicit inherited paint and line width agree with Chromium and PPTX', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="#f008" stroke="blue" stroke-width="4"><g fill="inherit" stroke="inherit" stroke-width="inherit"><circle cx="20" cy="20" r="10"/><g fill="none" stroke="red" stroke-width="2"><path d="M50 10 L70 20" fill="inherit" stroke="inherit" stroke-width="inherit" stroke-linecap="round" stroke-linejoin="round"/></g></g><circle cx="100" cy="20" r="10" fill="inherit" stroke="inherit" stroke-width="inherit"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const doc = createPresentation({size: {width: result.width, height: result.height}}); doc.addSlide({elements: [...result.elements]});
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((xml.match(/<a:ln w="38100"/g) ?? []).length, 2); assert.match(xml, /<a:ln w="19050" cap="rnd"/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const actual = await page.evaluate(source => { const host = document.createElement('div'); host.innerHTML = source; document.body.append(host); const values = [...host.querySelectorAll('circle,path')].map(node => { const style = getComputedStyle(node); return {fill: style.fill, stroke: style.stroke, width: style.strokeWidth}; }); host.remove(); return values; }, source);
  assert.deepEqual(actual.map(item => [item.stroke, item.width]), [['rgb(0, 0, 255)', '4px'], ['rgb(255, 0, 0)', '2px'], ['rgb(0, 0, 255)', '4px']]);
  assert.equal(actual[1].fill, 'none'); assert.equal(actual[0].fill, actual[2].fill);
});

test('bounded XML and external-resource rejection agree in Node and Chromium without execution or requests', async t => {
  const wrap = content => `<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">${content}</svg>`;
  const examples = [
    {name: 'valid', source: wrap('<rect x="10" y="10" width="10" height="10"/>'), expect: 'editable'},
    {name: 'malformed', source: wrap('<rect width="1" height="1">'), expect: 'invalid-xml'},
    {name: 'unknown entity', source: wrap('&unknown;'), expect: 'invalid-xml'},
    {name: 'internal entity', source: '<!DOCTYPE svg [<!ENTITY x "hello">]>' + wrap('&x;'), expect: 'invalid-xml'},
    {name: 'external DTD', source: '<!DOCTYPE svg SYSTEM "https://example.invalid/evil.dtd">' + wrap(''), expect: 'invalid-xml'},
    {name: 'foreign namespace', source: wrap('<rect xmlns="https://example.invalid/ns" width="10" height="10"/>'), expect: 'unsupported'},
    {name: 'prefixed namespace', source: '<s:svg xmlns:s="http://www.w3.org/2000/svg" width="100" height="100"/>', expect: 'unsupported'},
    {name: 'namespaced attribute', source: wrap('<rect xmlns:x="https://example.invalid/ns" x:width="10" height="10"/>'), expect: 'unsupported'},
    {name: 'remote image', source: wrap('<image href="https://example.invalid/image.png"/>'), expect: 'unsupported'},
    {name: 'remote fill', source: wrap('<rect width="10" height="10" fill="url(https://example.invalid/p)"/>'), expect: 'unsupported'},
    {name: 'script', source: wrap('<script>globalThis.__svgExecuted = true</script>'), expect: 'unsupported'},
    {name: 'processing instruction', source: '<?xml-stylesheet href="https://example.invalid/a.css"?>' + wrap(''), expect: 'unsupported'},
    {name: 'depth', source: wrap('<g>'.repeat(64) + '</g>'.repeat(64)), expect: 'input-limit'},
    {name: 'attributes', source: wrap(`<rect ${Array.from({length: 65}, (_, i) => `a${i}="x"`).join(' ')}/>`), expect: 'input-limit'},
    {name: 'element count', source: wrap('<rect width="1" height="1"/>'.repeat(10_000)), expect: 'input-limit'},
    {name: 'character count', source: ' '.repeat(1_000_001), expect: 'input-limit'},
  ];
  const attempt = source => { try { const result = importSimpleSvg(source); return {status: result.status, ...(result.status === 'unsupported' ? {issues: result.issues} : {elements: result.elements.length})}; } catch (error) { return {status: error.code, name: error.name, message: error.message}; } };
  const expected = examples.map(item => attempt(item.source));
  for (let i = 0; i < examples.length; i++) assert.equal(expected[i].status, examples[i].expect, examples[i].name);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const requests = []; page.on('request', request => requests.push(request.url()));
  await page.addScriptTag({content: await readFile('packages/svg-parser/dist/browser/global.js', 'utf8')});
  const observed = await page.evaluate(examples => {
    const outputs = examples.map(({source}) => { try { const result = PPTKit.importSimpleSvg(source); return {status: result.status, ...(result.status === 'unsupported' ? {issues: result.issues} : {elements: result.elements.length})}; } catch (error) { return {status: error.code, name: error.name, message: error.message}; } });
    return {outputs, executed: globalThis.__svgExecuted === true};
  }, examples);
  assert.deepEqual(observed.outputs, expected); assert.equal(observed.executed, false); assert.deepEqual(requests, []);
});

test('explicit fallback outcome preserves findings and embeds caller bytes in Node and Chromium without requests', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><image href="https://example.invalid/image.png"/></svg>';
  const bytes = new Uint8Array(await readFile('tests/fixtures/transparent-replacement.png'));
  const options = {sourceId: 'caller:source', fallback: {assetId: 'raster', bytes, mimeType: 'image/png', width: 150, height: 75}};
  const result = importSvgWithFallback(source, options); assert.equal(result.status, 'raster-fallback'); assert.equal(result.reason.status, 'unsupported');
  const doc = createPresentation({size: {width: 150, height: 75}}); doc.registerAsset(result.asset); doc.addSlide({elements: [result.element]});
  const exported = await generatePptx(doc), zip = readGeneratedZip(exported.bytes);
  assert.match(zip.get('ppt/slides/slide1.xml').toString(), /<p:pic>/);
  assert.deepEqual(new Uint8Array([...zip.entries()].find(([name]) => name.startsWith('ppt/media/'))[1]), bytes);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const requests = []; page.on('request', request => requests.push(request.url()));
  await page.addScriptTag({content: await readFile('packages/svg-parser/dist/browser/global.js', 'utf8')});
  const actual = await page.evaluate(({source, options}) => {
    options.fallback.bytes = new Uint8Array(options.fallback.bytes);
    const output = PPTKit.importSvgWithFallback(source, options);
    return {...output, asset: {...output.asset, source: {...output.asset.source, value: [...output.asset.source.value]}}};
  }, {source, options: {...options, fallback: {...options.fallback, bytes: [...bytes]}}});
  assert.deepEqual(actual, {...result, asset: {...result.asset, source: {...result.asset.source, value: [...result.asset.source.value]}}});
  assert.deepEqual(requests, []);
});

test('invalid caller raster is surfaced as exporter resource failure rather than an embedded image', async () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><image href="https://example.invalid/image.png"/></svg>';
  const result = importSvgWithFallback(source, {sourceId: 'bad-raster', fallback: {assetId: 'raster', bytes: new Uint8Array([1,2,3]), mimeType: 'image/png', width: 150, height: 75}});
  assert.equal(result.status, 'raster-fallback');
  const doc = createPresentation({size: {width: 150, height: 75}}); doc.registerAsset(result.asset); doc.addSlide({elements: [result.element]});
  const exported = await generatePptx(doc); assert.ok(exported.warnings.some(issue => issue.code === 'asset-read-failed')); assert.ok(exported.warnings.some(issue => issue.code === 'image-omitted'));
  assert.doesNotMatch(readGeneratedZip(exported.bytes).get('ppt/slides/slide1.xml').toString(), /<p:pic>/);
});

test('RGB and RGBA forms with inherited opacity match Chromium and PPTX', async t => {
  const colors = ['rgba(255,0,0,.5)', 'rgb(0 128 255 / 25%)', 'rgba(1 2 3 / 0)', 'rgb(10,20,30,1)', 'RGBA(4,5,6,50%)'];
  const source = `<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" fill-opacity=".5">${colors.map((color, i) => `<rect x="${i * 15}" y="5" width="10" height="10" fill="${color}"/>`).join('')}</svg>`;
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const doc = createPresentation({size: {width: result.width, height: result.height}}); doc.addSlide({elements: [...result.elements]});
  const preview = (await renderPresentationToSvg(doc)).slides[0].svg, xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  for (const alpha of ['25000', '12500', '0', '50000']) assert.ok(xml.includes(`<a:alpha val="${alpha}"/>`));
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const pixels = await page.evaluate(async ({source, preview}) => {
    const sample = async value => { const image = new Image(); image.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(value); await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 100; canvas.height = 100; const ctx = canvas.getContext('2d'); ctx.fillStyle = 'white'; ctx.fillRect(0,0,100,100); ctx.drawImage(image,0,0,100,100); return Array.from({length: 5}, (_, i) => [...ctx.getImageData(i * 15 + 5, 10, 1, 1).data]); };
    return {source: await sample(source), preview: await sample(preview)};
  }, {source, preview});
  for (let i = 0; i < colors.length; i++) for (let j = 0; j < 4; j++) assert.ok(Math.abs(pixels.source[i][j] - pixels.preview[i][j]) <= 1, `${colors[i]} channel ${j}`);
});

test('hex alpha fill and stroke match Chromium compositing and PPTX alpha values', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="#f008" stroke="#00ff0080" stroke-width="4"><circle cx="30" cy="30" r="20"/><g fill="transparent"><circle cx="90" cy="30" r="20" stroke="#00f8"/></g><rect x="130" y="10" width="20" height="20" stroke="none"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const doc = createPresentation({size: {width: result.width, height: result.height}}); doc.addSlide({elements: [...result.elements]});
  const preview = (await renderPresentationToSvg(doc)).slides[0].svg;
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:alpha val="53333"\/>/); assert.match(xml, /<a:alpha val="50196"\/>/); assert.match(xml, /<a:alpha val="0"\/>/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const pixels = await page.evaluate(async ({source, preview}) => {
    const sample = async value => { const image = new Image(); image.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(value); await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100; const ctx = canvas.getContext('2d'); ctx.fillStyle = 'white'; ctx.fillRect(0,0,200,100); ctx.drawImage(image,0,0,200,100); return [[30,30],[30,10],[90,30],[90,10],[140,20],[5,5]].map(([x,y]) => [...ctx.getImageData(x,y,1,1).data]); };
    return {source: await sample(source), preview: await sample(preview)};
  }, {source, preview});
  for (let i = 0; i < pixels.source.length; i++) for (let j = 0; j < 4; j++) assert.ok(Math.abs(pixels.source[i][j] - pixels.preview[i][j]) <= 1, `${i}/${j}`);
  assert.deepEqual(pixels.source[2], [255,255,255,255]);
});

test('inherited paint opacity matches Chromium compositing and PPTX quantization', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="#f008" fill-opacity="50%" stroke="#00ff0080" stroke-opacity=".8" stroke-width="4"><circle cx="30" cy="30" r="20"/><g fill-opacity=".25" stroke-opacity="25%"><circle cx="90" cy="30" r="20"/></g><circle cx="150" cy="30" r="20"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const doc = createPresentation({size: {width: result.width, height: result.height}}); doc.addSlide({elements: [...result.elements]});
  const preview = (await renderPresentationToSvg(doc)).slides[0].svg;
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:alpha val="26667"\/>/); assert.match(xml, /<a:alpha val="13333"\/>/); assert.match(xml, /<a:alpha val="40157"\/>/); assert.match(xml, /<a:alpha val="12549"\/>/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const pixels = await page.evaluate(async ({source, preview}) => {
    const sample = async value => { const image = new Image(); image.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(value); await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100; const ctx = canvas.getContext('2d'); ctx.fillStyle = 'white'; ctx.fillRect(0,0,200,100); ctx.drawImage(image,0,0,200,100); return [[30,30],[30,10],[90,30],[90,10],[150,30],[5,5]].map(([x,y]) => [...ctx.getImageData(x,y,1,1).data]); };
    return {source: await sample(source), preview: await sample(preview)};
  }, {source, preview});
  for (let i = 0; i < pixels.source.length; i++) for (let j = 0; j < 4; j++) assert.ok(Math.abs(pixels.source[i][j] - pixels.preview[i][j]) <= 1, `${i}/${j}`);
});

test('transparent linear stroke survives PPTX projection with alpha quantization', async () => {
  const result = importSimpleSvg('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" fill="none" stroke="#ff000080" stroke-linecap="round" stroke-linejoin="round"><path d="M10 10 L50 20"/></svg>');
  assert.equal(result.status, 'editable');
  const doc = createPresentation({size: {width: result.width, height: result.height}}); doc.addSlide({elements: [...result.elements]});
  let bytes = (await generatePptx(doc)).bytes;
  for (let i = 0; i < 3; i++) {
    const session = await importPptx(bytes, {projection: 'content'});
    try {
      const element = session.document.slides[0].elements[0]; assert.equal(element.type, 'connector'); assert.equal(element.pathMode, 'open-round');
      assert.equal(element.style.paint.color, 'FF0000'); assert.equal(element.style.paint.opacity, 0.50196);
      bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
  }
});

test('SVG polyline and polygon match browser samples and preserve editable PPTX geometry', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="none" stroke="blue" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20,20 50,20 50,50"/><g transform="translate(80,0)" stroke-linecap="butt"><polygon points="20,20 50,20 50,50 20,50"/></g></svg>';
  const input = importSimpleSvg(source); assert.equal(input.status, 'editable');
  const doc = createPresentation({size: {width: input.width, height: input.height}}); doc.addSlide({elements: [...input.elements]});
  const preview = (await renderPresentationToSvg(doc)).slides[0].svg, bytes = (await generatePptx(doc)).bytes;
  const imported = await importPptx(bytes, {projection: 'content'});
  try { assert.deepEqual(imported.document.slides[0].elements.map(item => item.pathMode), ['open-round', 'closed-round']); }
  finally { imported.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const pixels = await page.evaluate(async ({source, preview}) => {
    const sample = async value => { const image = new Image(); image.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(value); await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100; const ctx = canvas.getContext('2d'); ctx.fillStyle = 'white'; ctx.fillRect(0,0,200,100); ctx.drawImage(image,0,0,200,100); return [[20,20],[35,20],[50,35],[100,20],[115,20],[115,35],[5,5]].map(([x,y]) => [...ctx.getImageData(x,y,1,1).data]); };
    return { source: await sample(source), preview: await sample(preview) };
  }, {source, preview});
  assert.deepEqual(pixels.source, pixels.preview); assert.deepEqual(pixels.source[1], [0,0,255,255]); assert.deepEqual(pixels.source[5], [255,255,255,255]);
});

test('inherited path caps and joins match Chromium and connector output', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="none" stroke="red" stroke-linecap="round" stroke-linejoin="round"><path d="M20 20 L40 30 L60 20"/><g stroke-linecap="butt"><path d="M100 20 L120 30 L140 20" stroke-linecap="round"/></g></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const styles = await page.evaluate(source => { const host = document.createElement('div'); host.innerHTML = source; document.body.append(host); const values = [...host.querySelectorAll('path')].map(path => ({cap: getComputedStyle(path).strokeLinecap, join: getComputedStyle(path).strokeLinejoin})); host.remove(); return values; }, source);
  assert.deepEqual(styles, [{cap: 'round', join: 'round'}, {cap: 'round', join: 'round'}]);
  const doc = createPresentation({size: {width: result.width, height: result.height}}); doc.addSlide({elements: [...result.elements]});
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((xml.match(/cap="rnd"/g) ?? []).length, 2); assert.equal((xml.match(/<a:round\/>/g) ?? []).length, 2);
});

test('open and closed round paths survive three PPTX projection/export rounds', async () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="none" stroke="red" stroke-linecap="round" stroke-linejoin="round"><path d="M20 20 L40 30 L60 20"/><path d="M100 20 H140 V50 H100 Z"/></svg>';
  const input = importSimpleSvg(source);
  assert.equal(input.status, 'editable');
  const doc = createPresentation({size: {width: input.width, height: input.height}}); doc.addSlide({elements: [...input.elements]});
  let bytes = (await generatePptx(doc)).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, {projection: 'content'});
    try {
      const elements = session.document.slides[0].elements;
      assert.equal(elements.length, 2); assert.deepEqual(elements.map(element => element.pathMode), ['open-round', 'closed-round']);
      for (let i = 0; i < 2; i++) {
        assert.deepEqual(elements[i].start, input.elements[i].start); assert.deepEqual(elements[i].end, input.elements[i].end);
        assert.deepEqual(elements[i].route, input.elements[i].route);
      }
      const noOp = await session.prepareSave({mode: 'no-op'}); assert.equal(noOp.status, 'ready'); assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
      bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
  }
});

test('rotated and reflected round custom paths retain visual vertices after PPTX import', async t => {
  const transforms = [{rotation: 90}, {rotation: 35, flipH: true}, {rotation: 210, flipV: true}, {flipH: true, flipV: true}];
  const samples = [];
  for (const [index, transform] of transforms.entries()) {
    const doc = createPresentation({size: {width: 200, height: 100}});
    const closed = index % 2 === 1;
    doc.addSlide({elements: [{type: 'connector', start: {x: 30, y: 20}, route: closed ? [{x: 70, y: 20}, {x: 70, y: 50}] : [{x: 70, y: 20}], end: closed ? {x: 30, y: 20} : {x: 70, y: 50}, transform, pathMode: closed ? 'closed-round' : 'open-round', style: {paint: {type: 'solid', color: 'FF0000'}, width: 2}}]});
    const before = (await renderPresentationToSvg(doc)).slides[0].svg;
    const session = await importPptx((await generatePptx(doc)).bytes, {projection: 'content'});
    try {
      const element = session.document.slides[0].elements[0]; assert.equal(element.pathMode, closed ? 'closed-round' : 'open-round');
      samples.push({ before, after: (await renderPresentationToSvg(session.document)).slides[0].svg });
    } finally { session.dispose(); }
  }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const vertices = await page.evaluate(samples => samples.map(({before, after}) => [before, after].map(svg => {
    const host = document.createElement('div'); host.innerHTML = svg; document.body.append(host);
    const node = host.querySelector('polyline,polygon'), matrix = node.getCTM();
    const points = [...node.points].map(p => { const v = new DOMPoint(p.x, p.y).matrixTransform(matrix); return {x: v.x, y: v.y}; });
    host.remove(); return points;
  })), samples);
  for (let i = 0; i < vertices.length; i++) for (let p = 0; p < 3; p++) for (const axis of ['x', 'y']) {
    assert.ok(Math.abs(vertices[i][0][p][axis] - vertices[i][1][p][axis]) < .001, `${JSON.stringify(transforms[i])} ${p} ${axis}`);
  }
});

test('round paths retain local coordinates and transform inside an imported group', async () => {
  const doc = createPresentation({size: {width: 300, height: 200}});
  doc.addSlide({elements: [{type: 'group', box: {x: 40, y: 20, width: 160, height: 120}, coordinateSize: {width: 80, height: 60}, children: [{type: 'connector', start: {x: 10, y: 10}, route: [{x: 50, y: 10}, {x: 50, y: 40}], end: {x: 10, y: 10}, pathMode: 'closed-round', style: {paint: {type: 'solid', color: '0000FF'}, width: 2}}]}]});
  const session = await importPptx((await generatePptx(doc)).bytes, {projection: 'content'});
  try {
    const group = session.document.slides[0].elements[0]; assert.equal(group.type, 'group'); assert.equal(group.children.length, 1);
    const child = group.children[0]; assert.equal(child.pathMode, 'closed-round');
    assert.deepEqual(child.start, {x: 10, y: 10}); assert.deepEqual(child.route, [{x: 50, y: 10}, {x: 50, y: 40}]); assert.deepEqual(child.end, child.start);
    const before = (await renderPresentationToSvg(doc)).slides[0].svg, after = (await renderPresentationToSvg(session.document)).slides[0].svg;
    assert.match(before, /points="10,10 50,10 50,40"/); assert.match(after, /points="10,10 50,10 50,40"/);
  } finally { session.dispose(); }
});

test('slide copy fits closed round paths without losing closure or mutating the source', async () => {
  const source = createPresentation({size: {width: 400, height: 200}});
  const originalSlide = source.addSlide({id: 'source', elements: [{id: 'path', type: 'connector', start: {x: 40, y: 20}, route: [{x: 120, y: 20}, {x: 120, y: 80}], end: {x: 40, y: 20}, pathMode: 'closed-round', style: {paint: {type: 'solid', color: 'FF0000', opacity: .4}, width: 4}}]});
  const original = structuredClone(originalSlide.elements);
  const dest = createPresentation({size: {width: 200, height: 200}});
  const copy = dest.copySlideFrom(source, 'source', {layoutId: null, theme: 'destination', geometry: 'fit-explicit-geometry', customData: 'preserve', assets: {deduplicate: 'content', external: 'reject'}}).slide;
  const path = copy.elements[0]; assert.equal(path.pathMode, 'closed-round');
  assert.deepEqual(path.start, {x: 20, y: 60}); assert.deepEqual(path.route, [{x: 60, y: 60}, {x: 60, y: 90}]); assert.deepEqual(path.end, path.start);
  assert.deepEqual(path.style, original[0].style); assert.deepEqual(originalSlide.elements, original); assert.deepEqual(validatePresentation(dest), []);
  const selection = getSvgSlideSelectionGeometry(resolveLayout(dest), 0);
  assert.equal(hitTestSvgSelectionGeometry(selection, {x: 40, y: 75}, .1)?.elementId, path.id);
  assert.equal(hitTestSvgSelectionGeometry(selection, {x: 30, y: 85}, .1), undefined);
  const xml = readGeneratedZip((await generatePptx(dest)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:close\/>/); assert.match(xml, /cap="rnd"/); assert.match(xml, /<a:alpha val="40000"\/>/);
  const session = await importPptx((await generatePptx(dest)).bytes, {projection: 'content'});
  try { const restored = session.document.slides[0].elements[0]; assert.equal(restored.pathMode, 'closed-round'); assert.deepEqual(restored.start, path.start); assert.deepEqual(restored.route, path.route); assert.deepEqual(restored.end, path.end); }
  finally { session.dispose(); }
  copy.replaceElement(path.id, {...path, end: {x: path.end.x + 1, y: path.end.y}});
  assert.ok(validatePresentation(dest).some(issue => issue.code === 'invalid-connector-path-mode'));
  await assert.rejects(generatePptx(dest), /validation/i);
  copy.replaceElement(path.id, path); assert.deepEqual(validatePresentation(dest), []);
});

test('text-only preservation saves retain round path source; path edits remain blocked', async () => {
  const doc = createPresentation({size: {width: 200, height: 100}});
  doc.addSlide({elements: [
    {id: 'text', type: 'text', box: {x: 5, y: 5, width: 150, height: 20}, content: 'Original'},
    {id: 'path', type: 'connector', start: {x: 20, y: 40}, route: [{x: 80, y: 40}, {x: 80, y: 80}], end: {x: 20, y: 40}, pathMode: 'closed-round', style: {paint: {type: 'solid', color: 'FF0000', opacity: .5}, width: 3}},
  ]});
  let bytes = (await generatePptx(doc)).bytes;
  const geometry = readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString().match(/<a:custGeom>[\s\S]*?<\/a:custGeom>/)[0];
  for (let round = 1; round <= 3; round++) {
    const session = await importPptx(bytes, {projection: 'content', editSaveProfile: 'text-runs-v1'});
    try {
      const slide = session.document.slides[0], text = structuredClone(slide.elements.find(item => item.type === 'text'));
      const path = structuredClone(slide.elements.find(item => item.type === 'connector')); assert.equal(path.pathMode, 'closed-round');
      assert.ok(session.report.findings.some(item => item.feature === 'drawing.connector.mapped-fields' && item.elementId === path.id && item.operation === 'edit-save' && item.tier === 'unsupported'));
      const noOp = await session.prepareSave({mode: 'no-op'}); assert.equal(noOp.status, 'ready'); assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
      text.content[0].runs[0].text = `Edited ${round}`; slide.replaceElement(text.id, text);
      const save = await session.prepareSave({mode: 'edit'}); assert.equal(save.status, 'ready', JSON.stringify(save.report));
      const changed = await session.save(save.plan); bytes = changed.bytes;
      assert.equal(readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString().match(/<a:custGeom>[\s\S]*?<\/a:custGeom>/)[0], geometry);
      path.route[0].x += 1; slide.replaceElement(path.id, path);
      const blocked = await session.prepareSave({mode: 'edit'}); assert.equal(blocked.status, 'blocked');
    } finally { session.dispose(); }
  }
  const reopened = await importPptx(bytes, {projection: 'content'});
  try { assert.equal(reopened.document.slides[0].elements.find(item => item.type === 'text').content[0].runs[0].text, 'Edited 3'); assert.equal(reopened.document.slides[0].elements.find(item => item.type === 'connector').pathMode, 'closed-round'); }
  finally { reopened.dispose(); }
});

test('custom path import rejects unsupported join and does not silently create a connector', async () => {
  const doc = createPresentation({size: {width: 200, height: 100}});
  doc.addSlide({elements: [{type: 'connector', start: {x: 10,y: 10}, end: {x: 100,y: 50}, pathMode: 'open-round', style: {paint: {type: 'solid', color: 'FF0000'}, width: 2}}]});
  const zip = unzipSync((await generatePptx(doc)).bytes), xml = strFromU8(zip['ppt/slides/slide1.xml']);
  assert.match(xml, /<a:round\/>/); zip['ppt/slides/slide1.xml'] = strToU8(xml.replace('<a:round/>', '<a:bevel/>'));
  const session = await importPptx(zipSync(zip), {projection: 'content'});
  try { assert.equal(session.document.slides[0].elements.length, 0); assert.ok(session.report.findings.some(item => item.feature === 'drawing.sp')); }
  finally { session.dispose(); }
});

test('round connector path modes validate and export explicit caps and closure without changing legacy connectors', async () => {
  const doc = createPresentation({size: {width: 200, height: 100}});
  const slide = doc.addSlide({elements: [
    {type: 'connector', start: {x: 10, y: 10}, end: {x: 40, y: 20}, pathMode: 'open-round'},
    {type: 'connector', start: {x: 60, y: 10}, route: [{x: 90, y: 10}, {x: 90, y: 40}], end: {x: 60, y: 10}, pathMode: 'closed-round'},
    {type: 'connector', start: {x: 110, y: 10}, end: {x: 140, y: 20}},
  ]});
  assert.deepEqual(validatePresentation(doc), []);
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((xml.match(/cap="rnd"/g) ?? []).length, 2); assert.equal((xml.match(/<a:round\/>/g) ?? []).length, 2); assert.equal((xml.match(/<a:close\/>/g) ?? []).length, 1);
  const svg = (await renderPresentationToSvg(doc)).slides[0].svg; assert.match(svg, /<polygon /); assert.equal((svg.match(/<polyline /g) ?? []).length, 2);
  slide.addElement({type: 'connector', start: {x: 1, y: 1}, end: {x: 3, y: 3}, pathMode: 'closed-round', style: {endArrow: 'arrow'}});
  assert.ok(validatePresentation(doc).some(issue => issue.code === 'invalid-connector-path-mode'));
});

test('implicit linear path coordinates match Chromium path length and projected vertices', async t => {
  const paths = ['M20 20 40 20 40 40 20 40 Z', 'm20 20 10 0 0 10 h-5-5 v-5 z', 'M20,20 L40,30,60,20 H80,90 V40,50'];
  const cases = paths.map(data => { const source = `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><path d="${data}" fill="none" stroke="red" stroke-linecap="round" stroke-linejoin="round"/></svg>`; const result = importSimpleSvg(source); assert.equal(result.status, 'editable', data); return {source, connector: result.elements[0]}; });
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const lengths = await page.evaluate(cases => cases.map(({source}) => { const host = document.createElement('div'); host.innerHTML = source; document.body.append(host); const length = host.querySelector('path').getTotalLength(); host.remove(); return length; }), cases);
  for (let i = 0; i < cases.length; i++) {
    const c = cases[i].connector, points = [c.start, ...c.route, c.end];
    const length = points.slice(1).reduce((sum, point, j) => sum + Math.hypot(point.x - points[j].x, point.y - points[j].y), 0) / .75;
    assert.ok(Math.abs(length - lengths[i]) < .001, paths[i]);
  }
});

test('closed linear path connector agrees with native SVG and PPTX route geometry', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><g transform="translate(10,10) scale(2)" fill="none" stroke="red" stroke-width="2"><path d="M 10 10 l 10 0 v 10 H 10 z" stroke-linecap="round" stroke-linejoin="round"/></g></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const doc = createPresentation({ size: { width: result.width, height: result.height } }); doc.addSlide({ elements: [...result.elements] });
  const preview = (await renderPresentationToSvg(doc)).slides[0].svg;
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:custGeom>/); assert.equal((xml.match(/<a:lnTo>/g) ?? []).length, 3); assert.match(xml, /<a:close\/>/); assert.match(xml, /<a:ln w="38100" cap="rnd">/); assert.match(xml, /<a:round\/>/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const samples = await page.evaluate(async ({source, preview}) => {
    const pixels = async svg => { const img = new Image(); img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg); await img.decode(); const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100; const ctx = canvas.getContext('2d'); ctx.fillStyle = 'white'; ctx.fillRect(0,0,200,100); ctx.drawImage(img,0,0,200,100); return [[30,30],[40,30],[50,30],[50,40],[40,40],[25,25]].map(([x,y]) => [...ctx.getImageData(x,y,1,1).data]); };
    return {original: await pixels(source), preview: await pixels(preview)};
  }, {source, preview});
  assert.deepEqual(samples.preview, samples.original); assert.deepEqual(samples.original[1], [255,0,0,255]); assert.deepEqual(samples.original[4], [255,255,255,255]);
});

test('all 148 named and comma RGB colors agree with Chromium computed paint and PPTX', async t => {
  const colorTable = await readFile('packages/svg-parser/src/import/named-colors.ts', 'utf8');
  const names = [...colorTable.matchAll(/^  ([a-z]+): "[A-F0-9]{6}",$/gm)].map(match => match[1]);
  assert.equal(names.length, 148); assert.equal(new Set(names).size, 148); assert.ok(names.includes('rebeccapurple'));
  const colors = [...names, 'rgb(1, 2, 3)', 'rgb(0,128,255)'];
  const source = `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="110">${colors.map((color, i) => `<rect x="${(i % 20) * 10}" y="${Math.floor(i / 20) * 10}" width="8" height="8" fill="${color}"/>`).join('')}</svg>`;
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const doc = createPresentation({ size: { width: result.width, height: result.height } }); doc.addSlide({ elements: [...result.elements] });
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ content: await readFile('packages/svg-parser/dist/browser/global.js', 'utf8') });
  const actual = await page.evaluate(source => {
    const host = document.createElement('div'); host.innerHTML = source; document.body.append(host);
    const css = [...host.querySelectorAll('rect')].map(rect => getComputedStyle(rect).fill); host.remove();
    return { css, imported: PPTKit.importSimpleSvg(source) };
  }, source);
  assert.deepEqual(actual.imported, result);
  for (let i = 0; i < colors.length; i++) {
    const channels = actual.css[i].match(/\d+/g).map(Number), hex = channels.map(value => value.toString(16).padStart(2, '0')).join('').toUpperCase();
    assert.equal(result.elements[i].style.fill.color, hex, colors[i]); assert.ok(xml.includes(`val="${hex}"`), colors[i]);
  }
});

test('uniformly scaled ellipse stroke matches Chromium samples and PPTX line width', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><g transform="translate(30,20) scale(2)"><ellipse cx="15" cy="15" rx="10" ry="5" fill="none" stroke="#f00" stroke-width="2"/></g></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const doc = createPresentation({ size: { width: result.width, height: result.height } }); doc.addSlide({ elements: [...result.elements] });
  const preview = (await renderPresentationToSvg(doc)).slides[0].svg;
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:ln w="38100">/); assert.match(xml, /val="FF0000"/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const pixels = await page.evaluate(async ({ source, preview }) => {
    const sample = async svg => { const image = new Image(); image.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg); await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100; const ctx = canvas.getContext('2d'); ctx.fillStyle = 'white'; ctx.fillRect(0, 0, 200, 100); ctx.drawImage(image, 0, 0, 200, 100); return [[60,40],[60,39],[60,35],[60,50],[80,50],[85,50]].map(([x,y]) => [...ctx.getImageData(x,y,1,1).data]); };
    return { original: await sample(source), preview: await sample(preview) };
  }, { source, preview });
  assert.deepEqual(pixels.preview, pixels.original); assert.ok(pixels.original.some(([r,g,b]) => r === 255 && g === 0 && b === 0));
});

test('inherited strokes and overrides match browser SVG paint and exported lines', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" stroke="#f00" stroke-width="4"><g stroke="#0f0" stroke-width="2" transform="translate(20,0)"><circle cx="20" cy="20" r="8"/><g stroke="none"><circle cx="65" cy="20" r="8" stroke="#00f"/></g></g><circle cx="120" cy="20" r="8"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const doc = createPresentation({ size: { width: result.width, height: result.height } }); doc.addSlide({ elements: [...result.elements] });
  const preview = (await renderPresentationToSvg(doc)).slides[0].svg;
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((xml.match(/<a:ln w="19050">/g) ?? []).length, 2); assert.match(xml, /<a:ln w="38100">/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const actual = await page.evaluate(async ({source, preview}) => {
    const sample = async svg => { const image = new Image(); image.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg); await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100; const ctx = canvas.getContext('2d'); ctx.fillStyle = 'white'; ctx.fillRect(0, 0, 200, 100); ctx.drawImage(image, 0, 0, 200, 100); return [[40,12],[85,12],[120,11],[120,20],[5,5]].map(([x,y]) => [...ctx.getImageData(x,y,1,1).data]); };
    return { original: await sample(source), preview: await sample(preview) };
  }, {source, preview});
  assert.deepEqual(actual.preview, actual.original);
  assert.deepEqual(actual.original.slice(0, 3), [[0,255,0,255], [0,0,255,255], [255,0,0,255]]);
});

test('caller raster fallback embeds supplied image in PPTX and agrees in Chromium without network access', async t => {
  const bytes = new Uint8Array(await readFile('tests/fixtures/transparent-replacement.png'));
  const options = { sourceId: 'asset:original-svg', assetId: 'fallback', bytes, mimeType: 'image/png', width: 150, height: 75 };
  const fallback = createSvgRasterFallback(options);
  const doc = createPresentation({ size: { width: 150, height: 75 } });
  doc.registerAsset(fallback.asset); doc.addSlide({ elements: [fallback.element] });
  const zip = readGeneratedZip((await generatePptx(doc)).bytes);
  assert.match(zip.get('ppt/slides/slide1.xml').toString(), /<p:pic>/);
  const media = [...zip.entries()].filter(([name]) => name.startsWith('ppt/media/'));
  assert.equal(media.length, 1); assert.deepEqual(new Uint8Array(media[0][1]), bytes);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const requests = []; page.on('request', request => requests.push(request.url()));
  await page.addScriptTag({ content: await readFile('packages/svg-parser/dist/browser/global.js', 'utf8') });
  const actual = await page.evaluate(options => {
    const value = PPTKit.createSvgRasterFallback({ ...options, bytes: new Uint8Array(options.bytes) });
    return { ...value, asset: { ...value.asset, source: { ...value.asset.source, value: [...value.asset.source.value] } } };
  }, { ...options, bytes: [...bytes] });
  assert.deepEqual(actual, { ...fallback, asset: { ...fallback.asset, source: { ...fallback.asset.source, value: [...fallback.asset.source.value] } } });
  assert.deepEqual(requests, []);
});

test('experimental SVG import agrees in Node and Chromium and projects editable geometry', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><rect x="10" y="20" width="40" height="30" fill="#f00"/><circle cx="120" cy="40" r="20" fill="#00ff00"/></svg>';
  const input = importSimpleSvg(source); assert.equal(input.status, 'editable');
  const doc = createPresentation({ size: { width: input.width, height: input.height } }); doc.addSlide({ elements: [...input.elements] });
  const preview = (await renderPresentationToSvg(doc)).slides[0].svg;
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:off x="95250" y="190500"\/>/); assert.match(xml, /<a:prstGeom prst="rect"/);
  assert.match(xml, /<a:prstGeom prst="ellipse"/); assert.match(xml, /val="FF0000"/); assert.match(xml, /val="00FF00"/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const requests = []; page.on('request', request => requests.push(request.url()));
  await page.addScriptTag({ content: await readFile('packages/svg-parser/dist/browser/global.js', 'utf8') });
  const actual = await page.evaluate(({ source, preview }) => new Promise(async (resolve, reject) => {
    const result = PPTKit.importSimpleSvg(source);
    const rejected = PPTKit.importSimpleSvg('<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><image href="https://example.invalid/a"/></svg>');
    const pixel = async svg => {
      const image = new Image(); image.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
      await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100;
      const ctx = canvas.getContext('2d'); ctx.fillStyle = 'white'; ctx.fillRect(0, 0, 200, 100); ctx.drawImage(image, 0, 0, 200, 100);
      return [[20, 30], [120, 40], [80, 80]].map(([x, y]) => Array.from(ctx.getImageData(x, y, 1, 1).data));
    };
    try { resolve({ result, rejected, sourcePixels: await pixel(source), previewPixels: await pixel(preview) }); } catch (error) { reject(String(error)); }
  }), { source, preview });
  assert.deepEqual(actual.result, input); assert.equal(actual.rejected.status, 'unsupported'); assert.equal('elements' in actual.rejected, false);
  assert.deepEqual(actual.sourcePixels, [[255, 0, 0, 255], [0, 255, 0, 255], [255, 255, 255, 255]]);
  assert.deepEqual(actual.previewPixels, actual.sourcePixels); assert.deepEqual(requests, []);
});

test('viewBox meet and none match independent Chromium samples and PPTX coordinates', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ content: await readFile('packages/svg-parser/dist/browser/global.js', 'utf8') });
  for (const aspect of ['xMidYMid meet', 'none']) {
    const source = `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="-10 -20 100 100" preserveAspectRatio="${aspect}"><rect x="0" y="0" width="20" height="40" fill="#f00"/><circle cx="40" cy="30" r="10" fill="#00ff00"/></svg>`;
    const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
    const doc = createPresentation({ size: { width: result.width, height: result.height } }); doc.addSlide({ elements: [...result.elements] });
    const preview = (await renderPresentationToSvg(doc)).slides[0].svg;
    const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
    assert.match(xml, aspect === 'none' ? /<a:off x="190500" y="190500"\/>/ : /<a:off x="571500" y="190500"\/>/);
    const actual = await page.evaluate(async ({ source, preview }) => {
      const pixel = async svg => {
        const image = new Image(); image.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg); await image.decode();
        const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100; const ctx = canvas.getContext('2d');
        ctx.fillStyle = 'white'; ctx.fillRect(0, 0, 200, 100); ctx.drawImage(image, 0, 0, 200, 100);
        return [[30, 30], [70, 30], [100, 50], [120, 80]].map(([x, y]) => Array.from(ctx.getImageData(x, y, 1, 1).data));
      };
      return { browserResult: PPTKit.importSimpleSvg(source), sourcePixels: await pixel(source), previewPixels: await pixel(preview) };
    }, { source, preview });
    assert.deepEqual(actual.browserResult, result); assert.deepEqual(actual.previewPixels, actual.sourcePixels);
    assert.deepEqual(actual.sourcePixels, aspect === 'none'
      ? [[255, 0, 0, 255], [255, 255, 255, 255], [0, 255, 0, 255], [255, 255, 255, 255]]
      : [[255, 255, 255, 255], [255, 0, 0, 255], [0, 255, 0, 255], [255, 255, 255, 255]]);
  }
});

test('all aligned meet and slice mappings match Chromium SVG geometry in both viewport orientations', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const cases = [];
  for (const [width, height] of [[200, 100], [100, 200]]) for (const [xName, fx] of [['Min', 0], ['Mid', .5], ['Max', 1]]) for (const [yName, fy] of [['Min', 0], ['Mid', .5], ['Max', 1]]) for (const mode of ['meet', 'slice']) {
    const x = mode === 'slice' && width < height ? 20 + 50 * fx : 40;
    const y = mode === 'slice' && width > height ? 20 + 50 * fy : 40;
    const source = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 100 100" preserveAspectRatio="x${xName}Y${yName} ${mode}"><rect x="${x}" y="${y}" width="10" height="10" fill="#f00"/></svg>`;
    const result = importSimpleSvg(source); assert.equal(result.status, 'editable', source);
    cases.push({ source, box: result.elements[0].box });
  }
  const observed = await page.evaluate(cases => cases.map(({ source }) => {
    const host = document.createElement('div'); host.innerHTML = source; const svg = host.firstElementChild;
    svg.style.cssText = 'position:absolute;left:0;top:0'; document.body.append(svg);
    const outer = svg.getBoundingClientRect(), rect = svg.querySelector('rect').getBoundingClientRect();
    const result = { x: rect.left - outer.left, y: rect.top - outer.top, width: rect.width, height: rect.height }; svg.remove(); return result;
  }), cases);
  for (let i = 0; i < cases.length; i++) for (const field of ['x', 'y', 'width', 'height']) {
    assert.ok(Math.abs(cases[i].box[field] / .75 - observed[i][field]) < .001, `${field}: ${cases[i].source}`);
  }
});

test('ordered translation, scaling and reflections match Chromium geometry after viewBox mapping', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ content: await readFile('packages/svg-parser/dist/browser/global.js', 'utf8') });
  const shapes = [
    '<rect x="5" y="4" width="10" height="5" transform="translate(10,20) scale(2,3)"/>',
    '<rect x="5" y="4" width="10" height="5" transform="scale(2,3) translate(10,20)"/>',
    '<rect x="5" y="4" width="10" height="5" transform="scale(-1,1) translate(-40,0)"/>',
    '<ellipse cx="20" cy="20" rx="10" ry="5" transform="translate(20,10) scale(2,.5)"/>',
    '<circle cx="20" cy="20" r="10" transform="translate(80,0) scale(-1,2)"/>',
  ];
  const cases = shapes.map(shape => {
    const source = `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 100 100">${shape}</svg>`;
    const result = importSimpleSvg(source); assert.equal(result.status, 'editable', source); return { source, result };
  });
  const observed = await page.evaluate(cases => cases.map(({ source }) => {
    const host = document.createElement('div'); host.innerHTML = source; const svg = host.firstElementChild;
    svg.style.cssText = 'position:absolute;left:0;top:0'; document.body.append(svg);
    const outer = svg.getBoundingClientRect(), rect = svg.firstElementChild.getBoundingClientRect();
    const result = { x: rect.left - outer.left, y: rect.top - outer.top, width: rect.width, height: rect.height, imported: PPTKit.importSimpleSvg(source) };
    svg.remove(); return result;
  }), cases);
  for (let i = 0; i < cases.length; i++) {
    assert.deepEqual(observed[i].imported, cases[i].result);
    for (const field of ['x', 'y', 'width', 'height']) assert.ok(Math.abs(cases[i].result.elements[0].box[field] / .75 - observed[i][field]) < .001, `${field}: ${cases[i].source}`);
  }
  const result = cases[2].result, doc = createPresentation({ size: { width: result.width, height: result.height } }); doc.addSlide({ elements: [...result.elements] });
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:off x="714375" y="38100"\/>/); assert.match(xml, /<a:prstGeom prst="rect"/);
});

test('nested groups preserve inherited fills and geometry while reporting flattened structure', async t => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="#f00"><g fill="#0f0" transform="translate(40,20)"><g transform="scale(-1,2) translate(-30,0)"><rect x="5" y="5" width="10" height="10" transform="translate(2,1)"/><rect x="15" y="5" width="5" height="5" fill="none"/></g></g><rect x="2" y="2" width="4" height="4"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable'); assert.deepEqual(result.issues.map(issue => issue.code), ['flattened-groups']);
  const doc = createPresentation({ size: { width: result.width, height: result.height } }); doc.addSlide({ elements: [...result.elements] });
  const preview = (await renderPresentationToSvg(doc)).slides[0].svg;
  const xml = readGeneratedZip((await generatePptx(doc)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((xml.match(/<p:sp>/g) ?? []).length, 3); assert.match(xml, /val="00FF00"/); assert.match(xml, /val="FF0000"/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ content: await readFile('packages/svg-parser/dist/browser/global.js', 'utf8') });
  const observed = await page.evaluate(async ({ source, preview }) => {
    const host = document.createElement('div'); host.innerHTML = source; const svg = host.firstElementChild;
    svg.style.cssText = 'position:absolute;left:0;top:0'; document.body.append(svg);
    const outer = svg.getBoundingClientRect(), geometry = [...svg.querySelectorAll('rect')].map(node => {
      const rect = node.getBoundingClientRect(); return { x: rect.left - outer.left, y: rect.top - outer.top, width: rect.width, height: rect.height, fill: getComputedStyle(node).fill };
    }); svg.remove();
    const pixels = async value => {
      const image = new Image(); image.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(value); await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100; const ctx = canvas.getContext('2d');
      ctx.fillStyle = 'white'; ctx.fillRect(0, 0, 200, 100); ctx.drawImage(image, 0, 0, 200, 100);
      return [[60, 35], [52, 35], [4, 4]].map(([x, y]) => Array.from(ctx.getImageData(x, y, 1, 1).data));
    };
    return { geometry, browserResult: PPTKit.importSimpleSvg(source), sourcePixels: await pixels(source), previewPixels: await pixels(preview) };
  }, { source, preview });
  assert.deepEqual(observed.browserResult, result); assert.deepEqual(observed.geometry.map(item => item.fill), ['rgb(0, 255, 0)', 'none', 'rgb(255, 0, 0)']);
  for (let i = 0; i < 3; i++) for (const field of ['x', 'y', 'width', 'height']) assert.ok(Math.abs(result.elements[i].box[field] / .75 - observed.geometry[i][field]) < .001);
  assert.deepEqual(observed.sourcePixels, [[0, 255, 0, 255], [255, 255, 255, 255], [255, 0, 0, 255]]); assert.deepEqual(observed.previewPixels, observed.sourcePixels);
});
