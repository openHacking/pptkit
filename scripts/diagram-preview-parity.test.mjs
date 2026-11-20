import assert from "node:assert/strict";
import test from "node:test";
import { readFile, mkdir } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { importPptx, inspectPptx } from "../packages/pptx-parser/dist/index.js";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { diagramPreviewFixture, cacheTextTransform } from "../packages/pptx-parser/test/diagram-preview-fixture.mjs";
import { transform } from "../packages/pptx-parser/test/content-fixture.mjs";
import { picture, imageFixture, encodeImages } from "../packages/pptx-parser/test/image-fixture.mjs";
import { relationships } from "../packages/pptx-parser/test/opc-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const options = { projection: "content", diagramPreview: "cache", editSaveProfile: "text-runs-v3" };

test("audio authoring and static placeholder match public browser bundles", async t => {
  const input = { type: 'audio', id: 'speaker', audioAssetId: 'audio', box: { x: 10, y: 20, width: 100, height: 60 } };
  const build = core => { const p = core.createPresentation(); p.registerAsset({ id: 'audio', kind: 'audio', source: { type: 'bytes', value: Uint8Array.of(1, 2, 3) } }); p.addSlide({ id: 'slide', elements: [input] }); return p; };
  const document = build({ createPresentation }), expected = { normalized: normalizePresentation(document), rendered: await renderPresentationToSvg(document) };
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (const pkg of ['core', 'layout', 'svg-renderer', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(new URL(`../packages/${pkg}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const actual = await page.evaluate(async input => {
    const p = PPTKit.createPresentation(); p.registerAsset({ id: 'audio', kind: 'audio', source: { type: 'bytes', value: Uint8Array.of(1, 2, 3) } }); p.addSlide({ id: 'slide', elements: [input] });
    let error; try { await PPTKit.generatePptx(p); } catch (e) { error = e.message; }
    return { normalized: PPTKit.normalizePresentation(p), rendered: await PPTKit.renderPresentationToSvg(p), error };
  }, input);
  assert.deepEqual(actual.normalized, expected.normalized); assert.deepEqual(actual.rendered, expected.rendered); assert.match(actual.error, /requires audioMode wave-pcm-v1/);
});

test("PCM audio export matches browser global and Node without fetching", async t => {
  const wave = readGeneratedZip(await readFile(new URL('../tests/fixtures/pcm-wave-powerpoint.pptx', import.meta.url))).get('ppt/media/media1.wav');
  const input = { type: 'audio', id: 'speaker', audioAssetId: 'audio', box: { x: 10, y: 20, width: 64, height: 64 } };
  const p = createPresentation(); p.registerAsset({ id: 'audio', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: wave } }); p.addSlide({ id: 'slide', elements: [input] });
  const expected = await generatePptx(p, { audioMode: 'wave-pcm-v1' });
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const pkg of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(new URL(`../packages/${pkg}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const actual = await page.evaluate(async ({ input, wave }) => {
    const p = PPTKit.createPresentation(); p.registerAsset({ id: 'audio', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: Uint8Array.from(wave) } }); p.addSlide({ id: 'slide', elements: [input] });
    const result = await PPTKit.generatePptx(p, { audioMode: 'wave-pcm-v1' }); return { bytes: Array.from(result.bytes), warnings: result.warnings };
  }, { input, wave: [...wave] });
  assert.deepEqual(Uint8Array.from(actual.bytes), expected.bytes); assert.deepEqual(actual.warnings, expected.warnings); assert.deepEqual(requests, []);
});

test("browser audio source inventory agrees with Node and makes no requests", async t => {
  const xml = picture().replace('<p:nvPr/>', '<p:nvPr><a:audioFile r:link="sound"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><m:media xmlns:m="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="office"/></p:ext></p:extLst></p:nvPr>');
  const f = imageFixture(xml), wav = Buffer.alloc(46);
  wav.write('RIFF'); wav.writeUInt32LE(38, 4); wav.write('WAVEfmt ', 8); wav.writeUInt32LE(16, 16);
  wav.writeUInt16LE(1, 20); wav.writeUInt16LE(1, 22); wav.writeUInt32LE(8000, 24); wav.writeUInt32LE(8000, 28);
  wav.writeUInt16LE(1, 32); wav.writeUInt16LE(8, 34); wav.write('data', 36); wav.writeUInt32LE(2, 40);
  f.parts['media/tone.wav'] = wav; f.mediaTypes['media/tone.wav'] = 'audio/wav';
  f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), f.rel('sound', 'audio', '../media/tone.wav'), { Id: 'office', Type: 'http://schemas.microsoft.com/office/2007/relationships/media', Target: '../media/tone.wav' }]);
  const bytes = encodeImages(f), opts = { projection: 'content', audioSourceProfile: 'wave-pcm-v1' }, session = await importPptx(bytes, opts);
  t.after(() => session.dispose());
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(new URL('../packages/pptx-parser/dist/browser/global.js', import.meta.url), 'utf8') });
  const actual = await page.evaluate(async ({ bytes, opts }) => {
    const s = await globalThis.PPTKit.importPptx(Uint8Array.from(bytes), opts);
    try { return s.report.audioSources; } finally { s.dispose(); }
  }, { bytes: [...bytes], opts });
  assert.deepEqual(actual, session.report.audioSources); assert.equal(actual[0].status, 'wave-pcm'); assert.equal(actual[1].companionStatus, 'same-part'); assert.deepEqual(requests, []);
});

test("media placeholder browser rendering matches Node without network", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  const names = ["core", "layout", "pptx-parser", "svg-renderer"];
  const imports = Object.fromEntries(await Promise.all(names.map(async pkg => [`@pptkit/${pkg}`, `data:text/javascript;base64,${(await readFile(new URL(`../packages/${pkg}/dist/browser/index.js`, import.meta.url))).toString("base64")}`])));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  await page.evaluate(async () => { globalThis.parser = await import("@pptkit/pptx-parser"); globalThis.renderer = await import("@pptkit/svg-renderer"); });
  for (const strict of [false, true]) {
    const xml = picture({ embed: null }).replace('<p:nvPr/>', '<p:nvPr><a:audioFile r:link="external"/></p:nvPr>');
    const bytes = encodeImages(imageFixture(xml, strict)), options = { projection: "content", mediaPreview: "placeholder-v1" };
    const s = await importPptx(bytes, options);
    try {
      const expected = await renderPresentationToSvg(s.document);
      const actual = await page.evaluate(async ({ bytes, options }) => {
        const session = await globalThis.parser.importPptx(Uint8Array.from(bytes), options);
        try { return { rendered: await globalThis.renderer.renderPresentationToSvg(session.document), diagnostics: session.report.diagnostics }; } finally { session.dispose(); }
      }, { bytes: [...bytes], options });
      assert.deepEqual(actual.rendered, expected); assert.deepEqual(actual.diagnostics, s.report.diagnostics);
      await page.setContent(actual.rendered.slides[0].svg);
      assert.equal(await page.getByText("Media", { exact: true }).isVisible(), true);
      assert.equal(await page.locator("image,video,audio").count(), 0);
    } finally { s.dispose(); }
  }
  assert.deepEqual(requests, []);
});

test("browser public bundles render cached geometry/text/images with independent screen bounds and no network", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage({ viewport: { width: 900, height: 650 } });
  const requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  const names = ["core", "layout", "pptx-parser", "svg-renderer"];
  const imports = Object.fromEntries(await Promise.all(names.map(async pkg => [`@pptkit/${pkg}`, `data:text/javascript;base64,${(await readFile(new URL(`../packages/${pkg}/dist/browser/index.js`, import.meta.url))).toString("base64")}`])));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  await page.evaluate(async () => { globalThis.parser = await import("@pptkit/pptx-parser"); globalThis.renderer = await import("@pptkit/svg-renderer"); });
  await mkdir(new URL("../.tmp/diagram-preview/", import.meta.url), { recursive: true });
  for (const strict of [false, true]) {
    const f = diagramPreviewFixture(strict);
    f.setCache(f.cacheShape() + f.cacheShape("3", { text: "Image caption", xfrm: transform().replace('x="127000"', 'x="2286000"'), fill: '<a:blipFill><a:blip r:embed="cache-image"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>' }));
    const bytes = f.bytes(), s = await importPptx(bytes, options);
    try {
      const expected = await renderPresentationToSvg(s.document);
      const actual = await page.evaluate(async ({ bytes, options }) => {
        const s = await globalThis.parser.importPptx(Uint8Array.from(bytes), options);
        try { return { previews: s.report.diagramPreviews, rendered: await globalThis.renderer.renderPresentationToSvg(s.document) }; } finally { s.dispose(); }
      }, { bytes: [...bytes], options });
      assert.deepEqual(actual.previews, s.report.diagramPreviews); assert.deepEqual(actual.rendered, expected);
      await page.setContent(`<style>body{margin:0}</style>${actual.rendered.slides[0].svg}`);
      const boxes = await page.evaluate(() => {
        const shape = [...document.querySelectorAll('rect')].find(n => n.getAttribute('fill')?.toUpperCase() === '#228844');
        const image = document.querySelector('image');
        const box = n => { const b = n.getBoundingClientRect(); return { x: b.x, y: b.y, width: b.width, height: b.height }; };
        return { shape: box(shape), image: box(image), embedded: image.getAttribute('href')?.startsWith('data:image/png;base64,') ?? image.getAttributeNS('http://www.w3.org/1999/xlink', 'href')?.startsWith('data:image/png;base64,') };
      });
      assert.deepEqual(boxes.shape, { x: 110, y: 70, width: 100, height: 50 });
      assert.deepEqual(boxes.image, { x: 280, y: 70, width: 100, height: 50 }); assert.equal(boxes.embedded, true);
      assert.equal(await page.getByText("Cached text", { exact: true }).isVisible(), true);
      await page.screenshot({ path: new URL(`../.tmp/diagram-preview/${strict ? "strict" : "transitional"}.png`, import.meta.url).pathname });
    } finally { s.dispose(); }
  }
  assert.deepEqual(requests, []);
});

test("cached approximations survive three unrelated source edits; fresh output remains ordinary shapes", async () => {
  for (const strict of [false, true]) {
    const f = diagramPreviewFixture(strict);
    f.setCache(f.cacheShape("2", { xfrm: transform('rot="5400000"'), extra: cacheTextTransform('rot="-5400000"') }) + f.cacheShape("3", { text: "Image caption", xfrm: transform().replace('x="127000"', 'x="2286000"'), fill: '<a:blipFill><a:blip r:embed="cache-image"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>' }));
    let bytes = new Uint8Array(f.bytes());
    for (let round = 0; round < 3; round++) {
      const before = readGeneratedZip(bytes), s = await importPptx(bytes, options);
      try {
        const slide = s.document.slides[0], text = structuredClone(slide.elements[1]); text.content[0].runs[0].text = `Round ${round}`;
        slide.removeElement(text.id); slide.insertElement(1, text);
        const prepared = await s.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready"); bytes = (await s.save(prepared.plan)).bytes;
        const after = readGeneratedZip(bytes);
        for (const [name, body] of before) if (name !== "deck/z.xml") assert.deepEqual(after.get(name), body, name);
        const reopened = await importPptx(bytes, options);
        try { assert.deepEqual(reopened.report.diagramPreviews, s.report.diagramPreviews); const text = reopened.document.slides[0].elements[0].children[0].children[0].children[1]; assert.equal(text.content[0].runs[0].text, "Cached text"); assert.equal(text.transform.rotation, 0); } finally { reopened.dispose(); }
        const fresh = await generatePptx(s.document); assert.equal((await inspectPptx(fresh.bytes)).slides[0].diagrams, undefined);
      } finally { s.dispose(); }
    }
  }
});

test("independent cached text keeps its own screen rectangle while shape rotation is compensated", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 900, height: 650 } }); await page.route("**/*", route => route.abort());
  await mkdir(new URL("../.tmp/diagram-preview/", import.meta.url), { recursive: true });
  for (const strict of [false, true]) for (const rich of [false, true]) {
    const f = diagramPreviewFixture(strict);
    let child = f.cacheShape("2", { text: "AB", xfrm: transform('rot="5400000"'), extra: cacheTextTransform('rot="-5400000"') });
    if (rich) child = child.replace('</a:p>', '<a:r><a:rPr sz="1200"/><a:t> rich</a:t></a:r></a:p>');
    f.setCache('<dsp:grpSp><dsp:nvGrpSpPr/><dsp:grpSpPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="127000" y="254000"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></dsp:grpSpPr>' + child + '</dsp:grpSp>');
    const s = await importPptx(f.bytes(), options);
    try {
      await page.setContent(`<style>body{margin:0}</style>${(await renderPresentationToSvg(s.document)).slides[0].svg}`);
      const actual = await page.evaluate(rich => {
        const container = [...document.querySelectorAll('[data-pptkit-element-id]')].find(n => n.getAttribute('data-pptkit-element-id').endsWith('-text') && n.textContent === (rich ? 'AB rich' : 'AB'));
        const text = container.localName === 'foreignObject' ? container : container.querySelector(rich ? 'foreignObject' : 'text');
        const shape = [...document.querySelectorAll('rect')].find(n => n.getAttribute('fill')?.toUpperCase() === '#228844');
        const b = n => { const r = n.getBoundingClientRect(); return [r.x, r.y, r.width, r.height]; };
        const baseline = rich ? null : new DOMPoint(Number(text.getAttribute('x')), Number(text.getAttribute('y'))).matrixTransform(text.getCTM());
        return { text: b(text), shape: b(shape), baseline: baseline ? [baseline.x, baseline.y] : null, matrix: [text.getCTM().a, text.getCTM().b, text.getCTM().c, text.getCTM().d] };
      }, rich);
      if (rich) assert.deepEqual(actual.text, [150, 80, 160, 40]);
      else { assert.equal(actual.baseline[0], 152); assert.ok(Math.abs(actual.baseline[1] - 114.8) < 1e-6); }
      assert.deepEqual(actual.matrix, [2, 0, 0, 2]);
      assert.deepEqual(actual.shape, [160, 20, 100, 200]);
      await page.screenshot({ path: new URL(`../.tmp/diagram-preview/independent-${rich ? "rich" : "native"}-${strict ? "strict" : "transitional"}.png`, import.meta.url).pathname });
    } finally { s.dispose(); }
  }
});
