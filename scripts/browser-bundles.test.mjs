import assert from "node:assert/strict";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { chromium } from "@playwright/test";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { pack, basicParts, basicTypes, relationships, contentTypes } from "../packages/pptx-parser/test/opc-fixture.mjs";
import { structureFixture, encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { metadataFixture, encodeMetadata, custom } from "../packages/pptx-parser/test/metadata-fixture.mjs";
import { withSections, SECTION_IDS } from "../packages/pptx-parser/test/sections-fixture.mjs";
import { bytes as contentBytes, shape, body } from "../packages/pptx-parser/test/content-fixture.mjs";
import { png, picture, imageFixture, encodeImages } from "../packages/pptx-parser/test/image-fixture.mjs";
import { createImageImportFixture, createReferenceImportFixture, createNotesImportFixture, createTableImportFixture, createChartImportFixture } from "../tests/fixtures/content-import.mjs";
import { placeholderBytes } from "../packages/pptx-parser/test/placeholder-fixture.mjs";
import { textStyleFixture, setTextTheme, setLayoutThemeFont, layoutList } from "../packages/pptx-parser/test/text-style-fixture.mjs";
import { alternateParts, closureBytes } from "../packages/pptx-parser/test/closure-fixture.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const browserPackages = ["core", "layout", "pptx-exporter", "svg-renderer", "pptx-parser"];
function bundlePath(packageName, format) {
  return path.join(root, "packages", packageName, "dist", "browser", format === "esm" ? "index.js" : "global.js");
}

function javascriptDataUrl(source) {
  return `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;
}

test('fixed-axis value warnings agree in browser ESM and global consumers', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath('core', 'global'), 'utf8') });
  const url = javascriptDataUrl(await readFile(bundlePath('core', 'esm')));
  const results = await page.evaluate(async moduleUrl => {
    const esm = await import(moduleUrl), results = [];
    for (const api of [esm, globalThis.PPTKit]) {
      const deck = api.createPresentation();
      deck.addSlide({ id: 's', elements: [{ id: 'c', type: 'chart', chartType: 'line', box: { x: 0, y: 0, width: 300, height: 200 }, categories: ['A', 'B'], series: [{ name: 'S', values: [9, 37.25] }], axes: { value: { scale: { min: 0, max: 9, majorUnit: 1 } } } }] });
      deck.slides[0].addElement({ id: 'xy', type: 'chart', chartType: 'scatter', box: { x: 0, y: 0, width: 300, height: 200 }, series: [{ name: 'S', xValues: [11], yValues: [2] }], axes: { x: { scale: { min: 0, max: 9, majorUnit: 1 } } } });
      deck.slides[0].addElement({ id: 'combo', type: 'chart', chartType: 'combo', box: { x: 0, y: 0, width: 300, height: 200 }, categories: ['A'], series: [{ kind: 'bar', name: 'B', values: [5] }, { kind: 'line', name: 'L', values: [11] }], axes: { value: { scale: { min: 0, max: 9, majorUnit: 1 } }, secondaryValue: { scale: { min: 0, max: 10, majorUnit: 1 } } } });
      deck.slides[0].addElement({ id: 'stack', type: 'chart', chartType: 'bar', grouping: 'stacked', box: { x: 0, y: 0, width: 300, height: 200 }, categories: ['A'], series: [{ name: 'A', values: [6] }, { name: 'B', values: [6] }], axes: { value: { scale: { min: 0, max: 10, majorUnit: 1 } } } });
      results.push(api.validatePresentation(deck));
    }
    return results;
  }, url);
  assert.deepEqual(results[0], results[1]);
  assert.equal(results[0].length, 4); assert.ok(results[0].every(d => d.code === 'chart-value-outside-fixed-axis' && d.severity === 'warning'));
  assert.deepEqual(results[0].map(d => d.elementId), ['c', 'xy', 'combo', 'stack']);
  assert.match(results[0][0].path, /series\.0\.values\.1$/); assert.match(results[0][1].path, /series\.0\.xValues\.0$/); assert.match(results[0][2].path, /series\.1\.values\.0$/);
  assert.deepEqual(requests, []);
});

test('document slide links agree across ESM and global bundles', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), imports = {};
  for (const name of ['core', 'layout']) imports[`@pptkit/${name}`] = javascriptDataUrl(await readFile(bundlePath(name, 'esm')));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const moduleUrl = javascriptDataUrl(await readFile(bundlePath('svg-renderer', 'esm')));
  const results = await page.evaluate(async url => {
    const esm = await import(url), deck = PPTKit.createPresentation();
    deck.addSlide({ id: 'target' });
    deck.addSlide({ id: 'source', elements: [{ type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 20, height: 20 }, action: { type: 'slide', slideId: 'target' } }] });
    const layout = PPTKit.resolveLayout(deck), options = { slideActionMode: 'document-links', slideIndices: [1, 0] };
    return Promise.all([esm.renderLayoutToSvg(layout, options), PPTKit.renderLayoutToSvg(layout, options)]);
  }, moduleUrl);
  assert.deepEqual(results[0], results[1]);
  assert.match(results[0].slides[0].svg, /href="#pptkit-link-slide-0"/);
});

test('browser raster adapter agrees across ESM and global bundles', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  const imports = {};
  for (const name of ['core', 'layout']) imports[`@pptkit/${name}`] = javascriptDataUrl(await readFile(bundlePath(name, 'esm')));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const moduleUrl = javascriptDataUrl(await readFile(bundlePath('svg-renderer', 'esm')));
  const dataUrl = `data:image/png;base64,${(await readFile('tests/fixtures/transparent-replacement.png')).toString('base64')}`;
  const results = await page.evaluate(async ({ url, dataUrl }) => {
    const esm = await import(url), out = [];
    for (const api of [esm, globalThis.PPTKit]) {
      const d = PPTKit.createPresentation({ size: { width: 40, height: 20 } });
      d.registerAsset({ id: 'asset', kind: 'image', source: { type: 'url', value: dataUrl } });
      d.addSlide({ id: 's', elements: [{ type: 'shape', id: 'r', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 }, style: { fill: { type: 'solid', color: 'FF0000' } } }, { type: 'image', id: 'i', assetId: 'asset', box: { x: 10, y: 0, width: 10, height: 10 } }] });
      const result = await api.renderLayoutToBrowserRaster(PPTKit.resolveLayout(d), { format: 'png', backgroundMode: 'transparent', imageAnimationPolicy: 'require-static', isStaticImage: request => request.dataUrl === dataUrl });
      out.push({ ...result, slides: result.slides.map(slide => ({ ...slide, bytes: [...slide.bytes] })) });
    }
    return out;
  }, { url: moduleUrl, dataUrl });
  assert.deepEqual(results[0], results[1]);
  assert.deepEqual(results[0].slides[0].bytes.slice(0, 8), [137, 80, 78, 71, 13, 10, 26, 10]);
  assert.ok(results[0].warnings.some(warning => warning.code === 'raster-image-static-attested'));
});

test('JPEG opacity policy agrees across ESM and global browser bundles', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const imports = {};
  for (const name of ['core', 'layout']) imports[`@pptkit/${name}`] = javascriptDataUrl(await readFile(bundlePath(name, 'esm')));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const moduleUrl = javascriptDataUrl(await readFile(bundlePath('svg-renderer', 'esm')));
  const input = (await readFile('tests/fixtures/jpeg-12bit-progressive.jpg')).toString('base64');
  const results = await page.evaluate(async ({ url, input }) => {
    const esm = await import(url), out = [];
    for (const api of [esm, globalThis.PPTKit]) {
      const d = PPTKit.createPresentation({ size: { width: 6, height: 4 } }); d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: `data:image/jpeg;base64,${input}` } });
      d.addSlide({ id: 's', elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 6, height: 4 } }] });
      try { await api.renderLayoutToBrowserRaster(PPTKit.resolveLayout(d), { format: 'png', jpegDecodePolicy: 'require-opaque' }); out.push('accepted'); }
      catch (error) { out.push(error.code); }
    }
    return out;
  }, { url: moduleUrl, input });
  assert.deepEqual(results, ['raster-resource-unavailable', 'raster-resource-unavailable']);
});

test('strict raster font policy agrees across ESM and global bundles', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const imports = {};
  for (const name of ['core', 'layout']) imports[`@pptkit/${name}`] = javascriptDataUrl(await readFile(bundlePath(name, 'esm')));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const moduleUrl = javascriptDataUrl(await readFile(bundlePath('svg-renderer', 'esm')));
  const font = { family: 'Strict Test Face', weight: 400, italic: false, format: 'woff2', base64: (await readFile('packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2')).toString('base64') };
  const results = await page.evaluate(async ({ url, font }) => {
    const esm = await import(url), out = [];
    for (const api of [esm, globalThis.PPTKit]) {
      const d = PPTKit.createPresentation({ size: { width: 100, height: 40 } });
      d.addSlide({ id: 's', elements: [{ type: 'text', id: 't', box: { x: 0, y: 0, width: 90, height: 30 }, content: [{ runs: [{ text: 'Test', style: { fontFamily: font.family } }] }] }] });
      const layout = PPTKit.resolveLayout(d); let missing;
      try { await api.renderLayoutToBrowserRaster(layout, { format: 'png', fontPolicy: 'require-embedded' }); } catch (e) { missing = e.code; }
      const result = await api.renderLayoutToBrowserRaster(layout, { format: 'png', fontPolicy: 'require-embedded', embeddedFonts: [font] });
      const requests = [];
      const covered = await api.renderLayoutToBrowserRaster(layout, { format: 'png', fontPolicy: 'require-covered', embeddedFonts: [font], fontCoverage: request => { requests.push({ ...request }); return true; } });
      out.push({ missing, warnings: result.warnings, bytes: [...result.slides[0].bytes], covered: [...covered.slides[0].bytes], requests });
    }
    return out;
  }, { url: moduleUrl, font });
  assert.deepEqual(results[0], results[1]); assert.equal(results[0].missing, 'raster-font-unavailable');
});

test('embedded SVG font output agrees in Node and browser global', async t => {
  const { createPresentation } = await import('../packages/core/dist/index.js');
  const { renderPresentationToSvg } = await import('../packages/svg-renderer/dist/index.js');
  const slide = { id: 's', elements: [{ type: 'text', id: 't', box: { x: 0, y: 0, width: 120, height: 40 }, content: [{ runs: [{ text: 'Text', style: { fontFamily: 'Test Face' } }] }] }] };
  const font = { family: 'Test Face', weight: 400, italic: false, format: 'woff2', base64: (await readFile('packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2')).toString('base64') };
  const d = createPresentation(); d.addSlide(slide);
  const expected = await renderPresentationToSvg(d, { embeddedFonts: [font], portabilityDiagnostics: 'warn' });
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const actual = await page.evaluate(async ({ slide, font }) => {
    const d = PPTKit.createPresentation(); d.addSlide(slide);
    return PPTKit.renderPresentationToSvg(d, { embeddedFonts: [font], portabilityDiagnostics: 'warn' });
  }, { slide, font });
  assert.deepEqual(actual, expected);
});

test('native wrapped mixed SVG output matches Node and browser global', async t => {
  const { createPresentation } = await import('../packages/core/dist/index.js');
  const { renderPresentationToSvg } = await import('../packages/svg-renderer/dist/index.js');
  const slide = { id: 's', elements: [{ type: 'text', id: 't', box: { x: 0, y: 0, width: 65, height: 100 }, frame: { wrap: true }, content: [{ runs: [{ text: 'Alpha ' }, { text: 'Beta Gamma', style: { bold: true } }] }] }] };
  const d = createPresentation(); d.addSlide(slide);
  const options = { mixedTextMode: 'native', portabilityDiagnostics: 'warn' };
  const expected = await renderPresentationToSvg(d, options);
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const actual = await page.evaluate(async ({ slide, options }) => {
    const d = PPTKit.createPresentation(); d.addSlide(slide);
    return PPTKit.renderPresentationToSvg(d, options);
  }, { slide, options });
  assert.deepEqual(actual, expected);
  assert.doesNotMatch(actual.slides[0].svg, /<foreignObject/);
  assert.ok(actual.warnings.some(w => w.code === 'svg-native-mixed-text-approximation'));
});

test('native rotated table output matches Node and browser global', async t => {
  const { createPresentation } = await import('../packages/core/dist/index.js');
  const { renderPresentationToSvg } = await import('../packages/svg-renderer/dist/index.js');
  const slide = { id: 's', elements: [{ type: 'table', id: 't', box: { x: 0, y: 0, width: 100, height: 100 }, columns: [50, 50], rows: [{ height: 100, cells: [
    { content: 'CW', style: { textDirection: 'rotate90' } }, { content: 'CCW', style: { textDirection: 'rotate270' } },
  ] }] }] };
  const d = createPresentation(); d.addSlide(slide);
  const expected = await renderPresentationToSvg(d, { tableMode: 'native', portabilityDiagnostics: 'warn' });
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const actual = await page.evaluate(async slide => {
    const d = PPTKit.createPresentation(); d.addSlide(slide);
    return PPTKit.renderPresentationToSvg(d, { tableMode: 'native', portabilityDiagnostics: 'warn' });
  }, slide);
  assert.deepEqual(actual, expected);
  assert.ok(actual.warnings.some(w => w.code === 'table-text-direction-approximation'));
  assert.doesNotMatch(actual.slides[0].svg, /<foreignObject/);
});

test("SVG portability diagnostics match Node and browser global output", async t => {
  const { createPresentation } = await import('../packages/core/dist/index.js');
  const { renderPresentationToSvg } = await import('../packages/svg-renderer/dist/index.js');
  const slide = { id: 's', elements: [
      { type: 'text', id: 't', box: { x: 0, y: 0, width: 100, height: 30 }, frame: { wrap: false }, content: [{ runs: [{ text: 'A' }, { text: 'B', style: { bold: true } }] }] },
      { type: 'image', id: 'i', assetId: 'asset', box: { x: 0, y: 30, width: 20, height: 20 } },
      { type: 'shape', id: 'shape', shape: 'rect', box: { x: 30, y: 30, width: 100, height: 30 }, text: { frame: { wrap: false }, content: [{ runs: [{ text: 'C' }, { text: 'D', style: { italic: true } }] }] } },
      { type: 'table', id: 'table', box: { x: 0, y: 100, width: 100, height: 30 }, columns: [100], rows: [{ height: 30, cells: [{ content: 'Cell' }] }] },
    ] };
  const asset = { id: 'asset', kind: 'image', source: { type: 'url', value: 'https://example.invalid/a.png' } };
  const d = createPresentation(); d.registerAsset(asset); d.addSlide(slide);
  const expected = await renderPresentationToSvg(d, { portabilityDiagnostics: 'warn' });
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const actual = await page.evaluate(async ({ asset, slide }) => {
    const d = PPTKit.createPresentation(); d.registerAsset(asset); d.addSlide(slide);
    return PPTKit.renderPresentationToSvg(d, { portabilityDiagnostics: 'warn' });
  }, { asset, slide });
  assert.deepEqual(actual, expected);
  assert.deepEqual(actual.warnings.map(w => w.code), ['svg-foreign-object-dependency', 'svg-system-font-dependency', 'svg-external-image-dependency', 'svg-foreign-object-dependency', 'svg-system-font-dependency', 'svg-foreign-object-dependency', 'svg-system-font-dependency']);
  const nativeExpected = await renderPresentationToSvg(d, { portabilityDiagnostics: 'warn', mixedTextMode: 'native-no-wrap' });
  const nativeActual = await page.evaluate(async ({ asset, slide }) => {
    const d = PPTKit.createPresentation(); d.registerAsset(asset); d.addSlide(slide);
    return PPTKit.renderPresentationToSvg(d, { portabilityDiagnostics: 'warn', mixedTextMode: 'native-no-wrap' });
  }, { asset, slide });
  assert.deepEqual(nativeActual, nativeExpected);
  assert.deepEqual(nativeActual.warnings.map(w => w.code), ['svg-native-mixed-text-approximation', 'svg-system-font-dependency', 'svg-external-image-dependency', 'svg-native-mixed-text-approximation', 'svg-system-font-dependency', 'svg-foreign-object-dependency', 'svg-system-font-dependency']);
  const tableOptions = { portabilityDiagnostics: 'warn', mixedTextMode: 'native-no-wrap', tableMode: 'native' };
  const tableExpected = await renderPresentationToSvg(d, tableOptions);
  const tableActual = await page.evaluate(async ({ asset, slide, options }) => {
    const d = PPTKit.createPresentation(); d.registerAsset(asset); d.addSlide(slide);
    return PPTKit.renderPresentationToSvg(d, options);
  }, { asset, slide, options: tableOptions });
  assert.deepEqual(tableActual, tableExpected);
  assert.ok(tableActual.warnings.some(w => w.code === 'svg-native-table-approximation'));
  assert.doesNotMatch(tableActual.slides[0].svg, /<foreignObject/);
});

test("published browser artifacts are single-file and browser-safe", async () => {
  for (const packageName of browserPackages) {
    for (const format of ["esm", "global"]) {
      const file = bundlePath(packageName, format);
      assert.ok((await stat(file)).size > 0, `${file} should not be empty`);
      const source = await readFile(file, "utf8");
      assert.doesNotMatch(source, /(?:^|["'])node:/);
      assert.doesNotMatch(source, /\brequire\s*\(/);
      assert.doesNotMatch(source, /\bprocess(?:\.|\[)/);
      assert.doesNotMatch(source, /sourceMappingURL/);
      if (packageName === "pptx-parser") {
        assert.match(source, /Copyright \(c\) Contributors/);
        assert.match(source, /Copyright Louis-Dominique Dubeau/);
        assert.match(source, /saxes 6\.0\.0/);
      }
    }

    const packageRoot = path.join(root, "packages", packageName);
    const packed = spawnSync("npm", ["pack", "--dry-run", "--json"], {
      cwd: packageRoot,
      encoding: "utf8",
      env: { ...process.env, npm_config_cache: path.join("/tmp", "pptkit-npm-cache") },
    });
    assert.equal(packed.status, 0, packed.stderr);
    const [{ files }] = JSON.parse(packed.stdout);
    const names = new Set(files.map((file) => file.path));
    assert.ok(names.has("dist/browser/index.js"));
    assert.ok(names.has("dist/browser/global.js"));
    if (packageName === "pptx-parser") assert.ok(names.has("THIRD_PARTY_NOTICES.txt"));
  }

  const layoutEsm = await readFile(bundlePath("layout", "esm"), "utf8");
  const exporterEsm = await readFile(bundlePath("pptx-exporter", "esm"), "utf8");
  const rendererEsm = await readFile(bundlePath("svg-renderer", "esm"), "utf8");
  assert.match(layoutEsm, /from["']@pptkit\/core["']/);
  assert.match(exporterEsm, /from["']@pptkit\/core["']/);
  assert.match(exporterEsm, /from["']@pptkit\/layout["']/);
  assert.doesNotMatch(exporterEsm, /from["']fflate["']/);
  assert.match(rendererEsm, /from["']@pptkit\/core["']/);
  assert.match(rendererEsm, /from["']@pptkit\/layout["']/);
});

test("official ESM example previews and downloads a PPTX in Chromium", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  const cdnSources = {
    core: await readFile(bundlePath("core", "esm"), "utf8"),
    layout: await readFile(bundlePath("layout", "esm"), "utf8"),
    "pptx-exporter": await readFile(bundlePath("pptx-exporter", "esm"), "utf8"),
    "svg-renderer": await readFile(bundlePath("svg-renderer", "esm"), "utf8"),
  };
  let html = await readFile(path.join(root, "examples", "browser-esm", "index.html"), "utf8");
  for (const [packageName, source] of Object.entries(cdnSources)) {
    html = html.replace(
      `https://cdn.jsdelivr.net/npm/@pptkit/${packageName}@latest/dist/browser/index.js`,
      javascriptDataUrl(source),
    );
  }
  const appSource = await readFile(path.join(root, "examples", "browser-esm", "app.js"), "utf8");
  html = html.replace("./app.js", javascriptDataUrl(appSource));
  await page.setContent(html, { waitUntil: "load" });
  await page.evaluate(() => {
    const createObjectURL = URL.createObjectURL.bind(URL);
    URL.createObjectURL = (blob) => {
      globalThis.__pptkitBlobType = blob.type;
      return createObjectURL(blob);
    };
  });
  await page.getByRole("button", { name: "Preview SVG" }).click();
  await assert.doesNotReject(() => page.waitForFunction(() => document.querySelectorAll("#preview svg").length === 2));

  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download PPTX" }).click();
  const download = await downloadPromise;
  assert.equal(download.suggestedFilename(), "pptkit-browser-example.pptx");
  const downloadPath = await download.path();
  assert.ok(downloadPath !== null);
  const bytes = await readFile(downloadPath);
  assert.deepEqual(Array.from(bytes.subarray(0, 2)), [0x50, 0x4b]);
  assert.equal(await page.evaluate(() => globalThis.__pptkitBlobType), "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  await assert.doesNotReject(() => page.waitForFunction(() => document.querySelector("#status")?.textContent?.includes("Generated 2 slides")));
});

test("global bundles compose into one PPTKit namespace", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  for (const packageName of browserPackages) {
    await page.addScriptTag({ content: await readFile(bundlePath(packageName, "global"), "utf8") });
  }
  const result = await page.evaluate(async () => {
    const presentation = PPTKit.createPresentation();
    presentation.addSlide({ elements: [{ type: "text", content: "Global bundle", box: { x: 20, y: 20, width: 300, height: 60 } }] });
    presentation.addSlide();
    const [preview, generated] = await Promise.all([
      PPTKit.renderPresentationToSvg(presentation),
      PPTKit.generatePptx(presentation),
    ]);
    return {
      namespace: ["createPresentation", "resolveLayout", "generatePptx", "renderPresentationToSvg"].every((key) => typeof PPTKit[key] === "function"),
      previews: preview.slides.length,
      slideCount: generated.slideCount,
      signature: Array.from(generated.bytes.slice(0, 2)),
    };
  });
  assert.deepEqual(result, {
    namespace: true,
    previews: 2,
    slideCount: 2,
    signature: [0x50, 0x4b],
  });
});

test("GIF structural inspection matches browser global and Node", async (t) => {
  const { inspectGif } = await import("../packages/pptx-exporter/dist/index.js");
  const bytes = [...Buffer.from("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7", "base64")];
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const actual = await page.evaluate(input => PPTKit.inspectGif(new Uint8Array(input)), bytes);
  assert.deepEqual(actual, inspectGif(Uint8Array.from(bytes)));
});

test("GIF unknown and text block observations match Chromium and Node", async (t) => {
  const { inspectGif } = await import("../packages/pptx-exporter/dist/index.js");
  const bytes = [...Buffer.from("GIF89a"), 1, 0, 1, 0, 0x80, 0, 0, 0, 0, 0, 255, 255, 255,
    0x21, 1, 12, ...Array(12).fill(0), 1, 65, 0, 0x21, 0xee, 1, 42, 0, 0x21, 0xf9, 4, 0xe2, 7, 0, 0, 0,
    0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 1, 0x44, 0, 0x3b];
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const actual = await page.evaluate(input => PPTKit.inspectGif(new Uint8Array(input)), bytes);
  assert.deepEqual(actual, inspectGif(Uint8Array.from(bytes)));
  assert.equal(actual.plainTextBlocks, 1); assert.equal(actual.unknownExtensionBlocks, 1);
  assert.equal(actual.frames[0].userInputFlag, true); assert.equal(actual.frames[0].reservedControlBits, 0xe0);
});

test("non-8-bit JPEG preview warning agrees in Chromium global and Node", async t => {
  const { createPresentation } = await import('../packages/core/dist/index.js');
  const { renderPresentationToSvg } = await import('../packages/svg-renderer/dist/index.js');
  const bytes = await readFile('tests/fixtures/jpeg-12bit-progressive.jpg');
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  for (const [source, code] of [[bytes, 'jpeg-preview-precision-unverified'], [Buffer.concat([bytes.subarray(0, 2), Buffer.from([0xff, 0xe1, 0xff, 0xff]), Buffer.alloc(65533), bytes.subarray(2)]), 'jpeg-preview-inspection-incomplete']]) {
    const deck = createPresentation();
    deck.registerAsset({ id: 'jpeg', kind: 'image', source: { type: 'bytes', value: source } });
    deck.addSlide({ id: 'slide', elements: [{ type: 'image', id: 'image', assetId: 'jpeg', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    const expected = await renderPresentationToSvg(deck);
    const actual = await page.evaluate(async input => {
    const d = PPTKit.createPresentation();
    d.registerAsset({ id: 'jpeg', kind: 'image', source: { type: 'bytes', value: new Uint8Array(input) } });
    d.addSlide({ id: 'slide', elements: [{ type: 'image', id: 'image', assetId: 'jpeg', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    const result = await PPTKit.renderPresentationToSvg(d);
    return { warnings: result.warnings, svg: result.slides[0].svg, status: result.status };
    }, [...source]);
    assert.deepEqual(actual, { warnings: expected.warnings, svg: expected.slides[0].svg, status: expected.status });
    assert.deepEqual(actual.warnings.map(warning => warning.code), [code]);
  }
});

test("GIF LZW export validation matches Chromium global and Node", async (t) => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/node.js");
  const animation = "R0lGODlhAgACAIEAAP8AAAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwECAAAh+QQIAAAAACwAAAAAAgACAAAIBgABCAQQEAAh+QQIBwAAACwAAAAAAgACAIEAAP8AAAAAAAAAAAAIBgABCAQQEAA7";
  // This one-pixel image has a complete container, but its data block lacks an end code.
  const missingEnd = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
  const args = { animation, missingEnd };
  const build = async ({ animation, missingEnd, api = globalThis.PPTKit }) => {
    const deck = api.createPresentation();
    for (const [id, data] of Object.entries({ animation, missingEnd })) {
      deck.registerAsset({ id, kind: "image", source: { type: "url", value: `data:image/gif;base64,${data}` } });
    }
    deck.addSlide({ id: "slide", elements: ["animation", "missingEnd"].map((assetId, i) => ({
      type: "image", id: `image-${i}`, assetId, box: { x: i * 20, y: 0, width: 20, height: 20 },
    })) });
    const result = await api.generatePptx(deck, { externalImageValidation: "signature", gifDiagnostics: "reject-lzw" });
    return { bytes: [...result.bytes], warnings: result.warnings, status: result.status };
  };
  const expected = await build({ ...args, api: { createPresentation, generatePptx } });
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const actual = await page.evaluate(build, args);
  assert.deepEqual(actual, expected);
  assert.deepEqual(actual.warnings.map(w => [w.code, w.assetId]), [
    ["gif-animation-playback-unverified", "animation"], ["asset-read-failed", "missingEnd"], ["image-omitted", "missingEnd"],
  ]);
  assert.match(actual.warnings[1].message, /missing end code/);
});

test("single-frame GIF control warnings match Chromium global and Node", async (t) => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const gif = [...Buffer.from("GIF89a"), 1, 0, 1, 0, 0x80, 0, 0, 0, 0, 0, 255, 255, 255,
    0x21, 0xf9, 4, 0xe2, 0, 0, 0, 0, 0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 1, 0x44, 0, 0x3b];
  const deck = createPresentation();
  deck.registerAsset({ id: "image", kind: "image", source: { type: "bytes", value: Uint8Array.from(gif) } });
  deck.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "image", box: { x: 0, y: 0, width: 20, height: 20 } }] });
  const expected = await generatePptx(deck, { gifDiagnostics: "warn" });
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const actual = await page.evaluate(async input => {
    const d = PPTKit.createPresentation();
    d.registerAsset({ id: "image", kind: "image", source: { type: "bytes", value: new Uint8Array(input) } });
    d.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "image", box: { x: 0, y: 0, width: 20, height: 20 } }] });
    const result = await PPTKit.generatePptx(d, { gifDiagnostics: "warn" });
    return { bytes: [...result.bytes], warnings: result.warnings };
  }, gif);
  assert.deepEqual(actual, { bytes: [...expected.bytes], warnings: expected.warnings });
  assert.deepEqual(actual.warnings.map(w => w.code), ["gif-inspection-incomplete"]);
});

test("resolver animation-loss warnings match Chromium global and Node", async (t) => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const png = [...Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC", "base64")];
  const deck = createPresentation();
  deck.registerAsset({ id: "animation", kind: "image", source: { type: "path", value: "virtual" } });
  deck.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "animation", box: { x: 0, y: 0, width: 20, height: 20 } }] });
  const expected = await generatePptx(deck, { resolveAsset: () => ({ bytes: Uint8Array.from(png), animationFallback: "first-frame" }) });
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const actual = await page.evaluate(async input => {
    const d = PPTKit.createPresentation();
    d.registerAsset({ id: "animation", kind: "image", source: { type: "path", value: "virtual" } });
    d.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "animation", box: { x: 0, y: 0, width: 20, height: 20 } }] });
    const result = await PPTKit.generatePptx(d, { resolveAsset: () => ({ bytes: new Uint8Array(input), animationFallback: "first-frame" }) });
    return { bytes: [...result.bytes], warnings: result.warnings, status: result.status };
  }, png);
  assert.deepEqual(actual, { bytes: [...expected.bytes], warnings: expected.warnings, status: expected.status });
  assert.deepEqual(actual.warnings.map(w => w.code), ["asset-animation-flattened"]);
});

test("vector rasterization and failures match Chromium global and Node", async (t) => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const png = [...Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC", "base64")];
  const makeDeck = format => {
    const deck = createPresentation();
    deck.registerAsset({ id: "vector", kind: "image", mimeType: format === "svg" ? "image/svg+xml" : `image/${format}`, source: { type: "path", value: "virtual" } });
    deck.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "vector", box: { x: 2, y: 3, width: 30, height: 20 } }] });
    return deck;
  };
  const expected = [];
  for (const format of ["emf", "wmf", "svg"]) for (const limit of [undefined, 1]) {
    const result = await generatePptx(makeDeck(format), { ...(limit ? { maxAssetBytes: limit } : {}), resolveAsset: () => ({ bytes: Uint8Array.from(png), vectorRasterization: format }) });
    expected.push({ bytes: [...result.bytes], warnings: result.warnings, status: result.status });
  }
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const actual = await page.evaluate(async png => {
    const results = [];
    for (const format of ["emf", "wmf", "svg"]) for (const limit of [undefined, 1]) {
      const deck = PPTKit.createPresentation();
      deck.registerAsset({ id: "vector", kind: "image", mimeType: format === "svg" ? "image/svg+xml" : `image/${format}`, source: { type: "path", value: "virtual" } });
      deck.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "vector", box: { x: 2, y: 3, width: 30, height: 20 } }] });
      const result = await PPTKit.generatePptx(deck, { ...(limit ? { maxAssetBytes: limit } : {}), resolveAsset: () => ({ bytes: new Uint8Array(png), vectorRasterization: format }) });
      results.push({ bytes: [...result.bytes], warnings: result.warnings, status: result.status });
    }
    return results;
  }, png);
  assert.deepEqual(actual, expected);
  assert.deepEqual(actual.map(x => x.warnings[0].code), ["asset-vector-rasterized", "asset-read-failed", "asset-vector-rasterized", "asset-read-failed", "asset-vector-rasterized", "asset-read-failed"]);
});

test("paired SVG export matches Chromium and Node including aggregate budget failures", async t => {
  const { createPresentation } = await import('../packages/core/dist/index.js');
  const { generatePptx } = await import('../packages/pptx-exporter/dist/index.js');
  const png = [...Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC', 'base64')];
  const svg = [...new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>')];
  const deck = createPresentation(); deck.registerAsset({ id: 'svg', kind: 'image', mimeType: 'image/svg+xml', source: { type: 'bytes', value: Uint8Array.from(svg) } });
  deck.addSlide({ elements: [{ type: 'image', assetId: 'svg', box: { x: 0, y: 0, width: 20, height: 20 } }] });
  const expected = [];
  for (const cap of [png.length + svg.length, png.length + svg.length - 1]) {
    const result = await generatePptx(deck, { maxAssetBytes: cap, resolveAsset: () => ({ bytes: Uint8Array.from(png), svgPair: Uint8Array.from(svg) }) });
    expected.push({ bytes: [...result.bytes], warnings: result.warnings });
  }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (const name of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const actual = await page.evaluate(async ({ png, svg }) => {
    const deck = PPTKit.createPresentation(); deck.registerAsset({ id: 'svg', kind: 'image', mimeType: 'image/svg+xml', source: { type: 'bytes', value: Uint8Array.from(svg) } });
    deck.addSlide({ elements: [{ type: 'image', assetId: 'svg', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    const out = [];
    for (const cap of [png.length + svg.length, png.length + svg.length - 1]) {
      const result = await PPTKit.generatePptx(deck, { maxAssetBytes: cap, resolveAsset: () => ({ bytes: Uint8Array.from(png), svgPair: Uint8Array.from(svg) }) });
      out.push({ bytes: [...result.bytes], warnings: result.warnings });
    }
    return out;
  }, { png, svg });
  assert.deepEqual(actual, expected);
  assert.deepEqual(actual.map(x => x.warnings[0].code), ['asset-svg-pair-unverified', 'asset-read-failed']);
});

test("unsafe and case-colliding image IDs produce identical safe packages in Chromium", async t => {
  const { createPresentation } = await import('../packages/core/dist/index.js');
  const { generatePptx } = await import('../packages/pptx-exporter/dist/index.js');
  const bytes = [...png], ids = ['../slide', 'A', 'a', 'pptkit-asset-1'];
  const deck = createPresentation();
  for (const id of ids) deck.registerAsset({ id, kind: 'image', source: { type: 'bytes', value: Uint8Array.from(bytes) } });
  const slide = deck.addSlide(); for (const [i,id] of ids.entries()) slide.addElement({ type: 'image', assetId: id, box: { x: i * 20, y: 0, width: 20, height: 20 } });
  const expected = await generatePptx(deck);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (const name of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const actual = await page.evaluate(async ({ ids, bytes }) => {
    const deck = PPTKit.createPresentation();
    for (const id of ids) deck.registerAsset({ id, kind: 'image', source: { type: 'bytes', value: Uint8Array.from(bytes) } });
    const slide = deck.addSlide(); for (const [i,id] of ids.entries()) slide.addElement({ type: 'image', assetId: id, box: { x: i * 20, y: 0, width: 20, height: 20 } });
    const result = await PPTKit.generatePptx(deck); return { bytes: [...result.bytes], warnings: result.warnings };
  }, { ids, bytes });
  assert.deepEqual(actual, { bytes: [...expected.bytes], warnings: expected.warnings });
});

test("resolved-media deduplication matches Chromium and Node", async t => {
  const { createPresentation } = await import('../packages/core/dist/index.js');
  const { generatePptx } = await import('../packages/pptx-exporter/dist/index.js');
  const data = [...png], deck = createPresentation();
  deck.registerAsset({ id: 'a', kind: 'image', source: { type: 'bytes', value: Uint8Array.from(data) } });
  deck.registerAsset({ id: 'b', kind: 'image', source: { type: 'path', value: 'virtual' } });
  deck.addSlide({ elements: ['a','b'].map((assetId, i) => ({ type: 'image', assetId, box: { x: i * 20, y: 0, width: 20, height: 20 } })) });
  const expected = await generatePptx(deck, { deduplicateResolvedAssets: true, resolveAsset: () => ({ bytes: Uint8Array.from(data) }) });
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (const name of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const actual = await page.evaluate(async data => {
    const deck = PPTKit.createPresentation();
    deck.registerAsset({ id: 'a', kind: 'image', source: { type: 'bytes', value: Uint8Array.from(data) } });
    deck.registerAsset({ id: 'b', kind: 'image', source: { type: 'path', value: 'virtual' } });
    deck.addSlide({ elements: ['a','b'].map((assetId, i) => ({ type: 'image', assetId, box: { x: i * 20, y: 0, width: 20, height: 20 } })) });
    const result = await PPTKit.generatePptx(deck, { deduplicateResolvedAssets: true, resolveAsset: () => ({ bytes: Uint8Array.from(data) }) });
    return { bytes: [...result.bytes], warnings: result.warnings };
  }, data);
  assert.deepEqual(actual, { bytes: [...expected.bytes], warnings: expected.warnings });
  assert.equal([...readGeneratedZip(expected.bytes).keys()].filter(name => name.startsWith('ppt/media/')).length, 1);
});

test("caller rasterizes a local SVG in Chromium and exports its actual PNG", async (t) => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  let requests = 0; await page.route("**/*", route => { requests++; return route.abort(); });
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const result = await page.evaluate(async () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="10"><rect width="10" height="10" fill="red"/><rect x="10" width="10" height="10" fill="blue"/></svg>';
    const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
    const img = new Image(); await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = reject; img.src = url; });
    const canvas = document.createElement('canvas'); canvas.width = 20; canvas.height = 10;
    const ctx = canvas.getContext('2d'); ctx.drawImage(img, 0, 0); URL.revokeObjectURL(url);
    const samples = [[5,5],[15,5]].map(([x,y]) => [...ctx.getImageData(x,y,1,1).data]);
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
    const png = new Uint8Array(await blob.arrayBuffer());
    const deck = PPTKit.createPresentation();
    deck.registerAsset({ id: 'svg', kind: 'image', mimeType: 'image/svg+xml', source: { type: 'bytes', value: new TextEncoder().encode(svg) } });
    deck.addSlide({ elements: [{ type: 'image', assetId: 'svg', box: { x: 0, y: 0, width: 20, height: 10 } }] });
    const exported = await PPTKit.generatePptx(deck, { resolveAsset: () => ({ bytes: png, vectorRasterization: 'svg' }) });
    return { samples, png: [...png], bytes: [...exported.bytes], warnings: exported.warnings.map(w => w.code) };
  });
  assert.deepEqual(result.samples, [[255,0,0,255],[0,0,255,255]]);
  assert.deepEqual(result.warnings, ['asset-vector-rasterized']);
  assert.deepEqual(readGeneratedZip(Uint8Array.from(result.bytes)).get('ppt/media/svg.png'), Buffer.from(result.png));
  assert.equal(requests, 0);
});

test("JPEG ICC inventory and export diagnostics match Chromium and Node", async t => {
  const { createPresentation } = await import('../packages/core/dist/index.js');
  const { inspectJpegIcc, extractJpegIccProfile, generatePptx } = await import('../packages/pptx-exporter/dist/index.js');
  const data = [...await readFile(path.join(root, 'tests/fixtures/jpeg-icc-srgb.jpg'))];
  const deck = createPresentation(); deck.registerAsset({ id: 'jpeg', kind: 'image', source: { type: 'bytes', value: Uint8Array.from(data) } });
  deck.addSlide({ elements: [{ type: 'image', assetId: 'jpeg', box: { x: 0, y: 0, width: 20, height: 20 } }] });
  const expected = await generatePptx(deck, { jpegIccDiagnostics: 'warn' });
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (const name of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const actual = await page.evaluate(async data => {
    const bytes = Uint8Array.from(data), deck = PPTKit.createPresentation();
    deck.registerAsset({ id: 'jpeg', kind: 'image', source: { type: 'bytes', value: bytes } });
    deck.addSlide({ elements: [{ type: 'image', assetId: 'jpeg', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    const result = await PPTKit.generatePptx(deck, { jpegIccDiagnostics: 'warn' });
    let limit; try { PPTKit.inspectJpegIcc(bytes, { maxBytes: 1 }); } catch (e) { limit = e.name; }
    const profile = PPTKit.extractJpegIccProfile(bytes).profileBytes;
    const chunk = (index, payload) => { const body = [...new TextEncoder().encode('ICC_PROFILE\0'), index, 2, ...payload], len = body.length + 2; return [255,226,len >>> 8,len & 255,...body]; };
    const shuffled = Uint8Array.from([255,216,...chunk(2,profile.subarray(200)),...chunk(1,profile.subarray(0,200)),255,217]);
    let incomplete; try { PPTKit.extractJpegIccProfile(Uint8Array.from([255,216,...chunk(1,profile.subarray(0,200)),255,217])); } catch(e) { incomplete = e.message; }
    return { info: PPTKit.inspectJpegIcc(bytes), profile: [...profile], shuffled: [...PPTKit.extractJpegIccProfile(shuffled).profileBytes], incomplete, bytes: [...result.bytes], warnings: result.warnings, limit };
  }, data);
  assert.deepEqual(actual.info, inspectJpegIcc(Uint8Array.from(data)));
  assert.deepEqual(actual.profile, [...extractJpegIccProfile(Uint8Array.from(data)).profileBytes]); assert.deepEqual(actual.shuffled, actual.profile); assert.match(actual.incomplete, /incomplete/);
  assert.deepEqual(actual.bytes, [...expected.bytes]); assert.deepEqual(actual.warnings, expected.warnings); assert.equal(actual.limit, 'RangeError');
});

test("JPEG EXIF orientation inspection matches Chromium global and Node", async (t) => {
  const { inspectJpegExifOrientation } = await import("../packages/pptx-exporter/dist/index.js");
  const bytes = [...await readFile(path.join(root, "tests/fixtures/exif-orientation6.jpg"))];
  const expected = inspectJpegExifOrientation(Uint8Array.from(bytes));
  assert.deepEqual(expected, { status: "orientation", orientation: 6, stoppedAt: "scan" });
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const actual = await page.evaluate(input => PPTKit.inspectJpegExifOrientation(Uint8Array.from(input)), bytes);
  assert.deepEqual(actual, expected);
});

test("JPEG orientation export diagnostics match Chromium global and Node", async (t) => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const jpeg = [...await readFile(path.join(root, "tests/fixtures/exif-orientation6.jpg"))];
  const deck = createPresentation();
  deck.registerAsset({ id: "photo", kind: "image", source: { type: "bytes", value: Uint8Array.from(jpeg) } });
  deck.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "photo", box: { x: 0, y: 0, width: 30, height: 20 } }] });
  const expected = await generatePptx(deck, { jpegOrientationDiagnostics: "warn" });
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const actual = await page.evaluate(async input => {
    const deck = PPTKit.createPresentation();
    deck.registerAsset({ id: "photo", kind: "image", source: { type: "bytes", value: new Uint8Array(input) } });
    deck.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "photo", box: { x: 0, y: 0, width: 30, height: 20 } }] });
    const result = await PPTKit.generatePptx(deck, { jpegOrientationDiagnostics: "warn" });
    return { bytes: [...result.bytes], warnings: result.warnings, status: result.status };
  }, jpeg);
  assert.deepEqual(actual, { bytes: [...expected.bytes], warnings: expected.warnings, status: expected.status });
  assert.deepEqual(actual.warnings.map(w => w.code), ["jpeg-exif-orientation-unverified"]);
});

test("resolver orientation normalization matches Chromium global and Node", async (t) => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const jpg = [...await readFile(path.join(root, "tests/fixtures/exif-orientation6.jpg"))];
  const png = [...await readFile(path.join(root, "tests/fixtures/exif-orientation6-normalized.png"))];
  const deck = createPresentation();
  deck.registerAsset({ id: "photo", kind: "image", mimeType: "image/jpeg", source: { type: "bytes", value: Uint8Array.from(jpg) } });
  deck.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "photo", box: { x: 0, y: 0, width: 30, height: 20 } }] });
  const expected = await generatePptx(deck, { resolveAsset: () => ({ bytes: Uint8Array.from(png), mimeType: "image/png", orientationNormalization: "pixels" }) });
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const actual = await page.evaluate(async ({jpg,png}) => {
    const deck = PPTKit.createPresentation();
    deck.registerAsset({ id: "photo", kind: "image", mimeType: "image/jpeg", source: { type: "bytes", value: new Uint8Array(jpg) } });
    deck.addSlide({ id: "slide", elements: [{ type: "image", id: "image", assetId: "photo", box: { x: 0, y: 0, width: 30, height: 20 } }] });
    const result = await PPTKit.generatePptx(deck, { resolveAsset: () => ({ bytes: new Uint8Array(png), mimeType: "image/png", orientationNormalization: "pixels" }) });
    return { bytes: [...result.bytes], warnings: result.warnings, status: result.status };
  }, {jpg,png});
  assert.deepEqual(actual, { bytes: [...expected.bytes], warnings: expected.warnings, status: expected.status });
  assert.deepEqual(actual.warnings.map(w => w.code), ["asset-orientation-normalized"]);
});

test("XML inspector matches independent Chromium parsing and enforces its browser limits", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  let requests = 0;
  await page.route("**/*", (route) => { requests++; return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath("pptx-parser", "global"), "utf8") });
  const result = await page.evaluate(async () => {
    const cases = ['<r/>', '<p:r xmlns:p="urn:r" a="1">A&amp;B<p:c/></p:r>', '<r xmlns="urn:r"><c xmlns=""/></r>', '<r>😀<![CDATA[中]]><!--x--><?p y?></r>', '<r>&#xD800;</r>', '<r>\u0001</r>', '<r/>tail', '<r a="1" a="2"/>', '<r xmlns:a="urn:a" xmlns:b="urn:a" a:x="1" b:x="2"/>', '<r><x/></bad>', '<p:r/>'];
    const findings = [];
    for (const xml of cases) {
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      const domValid = !doc.querySelector("parsererror");
      try {
        const inspected = await PPTKit.inspectXmlPart(new TextEncoder().encode(xml));
        findings.push({ accepted: true, domValid, nameMatches: inspected.root.name === doc.documentElement.nodeName, uriMatches: inspected.root.namespaceUri === (doc.documentElement.namespaceURI ?? ""), elementsMatch: inspected.elements === doc.getElementsByTagName("*").length });
      } catch (error) { findings.push({ accepted: false, domValid, code: error.code }); }
    }
    const errors = [];
    for (const [xml, options] of [
      ['<!DOCTYPE r SYSTEM "https://example.invalid/external"><r/>', {}],
      ['<r><c/></r>', { limits: { maxDepth: 1 } }],
      ['<r>&entity;</r>', {}],
    ]) {
      try { await PPTKit.inspectXmlPart(new TextEncoder().encode(xml), options); errors.push("accepted"); }
      catch (error) { errors.push(error.code); }
    }
    return { findings, errors };
  });
  for (const finding of result.findings) {
    assert.equal(finding.accepted, finding.domValid);
    if (finding.accepted) assert.ok(finding.nameMatches && finding.uriMatches && finding.elementsMatch);
    else assert.equal(finding.code, "xml-invalid");
  }
  assert.deepEqual(result.errors, ["xml-unsupported", "xml-limit-exceeded", "xml-invalid"]);
  assert.equal(requests, 0);
});


test("OPC browser inventory resolves package references and rejects dangling targets without retrieval", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  let requests = 0;
  await page.route("**/*", (route) => { requests++; return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath("pptx-parser", "global"), "utf8") });
  const valid = pack({ ...basicParts(), "doc/_rels/main.xml.rels": relationships([{ Id: "external", Type: "urn:external", Target: "https://example.invalid/image.png", TargetMode: "External" }]) }, basicTypes());
  const broken = pack({ ...basicParts(), "_rels/.rels": relationships([{ Id: "bad", Type: "urn:root", Target: "absent.xml" }]) }, basicTypes());
  const result = await page.evaluate(async ({ valid, broken }) => {
    const result = await PPTKit.inspectOpcPackage(new Uint8Array(valid));
    const errors = [];
    for (const [bytes, options] of [[broken, {}], [valid, { limits: { maxXmlNodes: 1 } }]]) {
      try { await PPTKit.inspectOpcPackage(new Uint8Array(bytes), options); errors.push("accepted"); }
      catch (error) { errors.push(error.code); }
    }
    return { parts: result.parts.length, target: result.relationships[0].targetPart, external: result.relationships[1].target, errors };
  }, { valid: [...valid], broken: [...broken] });
  assert.deepEqual(result, { parts: 4, target: "/doc/main.xml", external: "https://example.invalid/image.png", errors: ["package-invalid", "package-limit-exceeded"] });
  assert.equal(requests, 0);
});

test("OPC comparison ESM/global bundles agree with independent DOM structure without resource retrieval", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const cases = [
    ['<r b="2" a="A&amp;B"><e/></r>', "<r a='A&#38;B' b='2'><e></e></r>", true],
    ['<r>A&amp;B😀</r>', '<r><![CDATA[A&B]]>&#x1F600;</r>', true],
    ['<r xmlns:q="urn:first" value="q:x"/>', '<r xmlns:q="urn:second" value="q:x"/>', false],
    ['<!--keep--><r/><?keep yes?>', '<r/><?keep yes?>', false],
    ['<r><a/><b/></r>', '<r><b/><a/></r>', false],
    ['<r><a/></r>', '<r> <a/></r>', false],
    ['<r xmlns:x="urn:unknown"><x:opaque v="one"/></r>', '<r xmlns:x="urn:unknown"><x:opaque v="two"/></r>', false],
  ];
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(); let requests = 0;
    await page.route("**/*", (route) => { requests++; return route.abort(); });
    const source = await readFile(bundlePath("pptx-parser", format), "utf8");
    if (format === "global") await page.addScriptTag({ content: source });
    const result = await page.evaluate(async ({ moduleUrl, fixtures }) => {
      const parser = moduleUrl ? await import(moduleUrl) : PPTKit;
      // Browser DOM is independent of Parser's saxes tree/token implementation.
      const canonical = (xml) => {
        const document = new DOMParser().parseFromString(xml, "application/xml");
        if (document.querySelector("parsererror")) throw new Error("Invalid oracle fixture");
        const children = (parent) => {
          const result = []; let text = "";
          const flush = () => { if (text) result.push(["text", text]); text = ""; };
          for (const node of parent.childNodes) {
            if (node.nodeType === 3 || node.nodeType === 4) { text += node.nodeValue; continue; }
            flush();
            if (node.nodeType === 1) result.push(["element", node.nodeName, node.namespaceURI ?? "", [...node.attributes].map((a) => [a.name, a.namespaceURI ?? "", a.value]).sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0), children(node)]);
            else if (node.nodeType === 8) result.push(["comment", node.nodeValue]);
            else if (node.nodeType === 7) result.push(["pi", node.nodeName, node.nodeValue]);
          }
          flush(); return result;
        };
        return JSON.stringify(children(document));
      };
      const results = [];
      for (const [a, b, expected, before, after] of fixtures) {
        const compared = await parser.compareOpcPackages(new Uint8Array(before), new Uint8Array(after), { profile: "opc-xml-v1" });
        const strict = await parser.compareOpcPackages(new Uint8Array(before), new Uint8Array(after));
        results.push({ equal: compared.equal, oracle: canonical(a) === canonical(b), expected, strict: strict.equal, frozen: Object.isFrozen(compared.differences) });
      }
      const controller = new AbortController(); controller.abort();
      let aborted;
      try { await parser.compareOpcPackages(new Uint8Array(fixtures[0][3]), new Uint8Array(fixtures[0][4]), { signal: controller.signal }); }
      catch (error) { aborted = error.code; }
      return { results, aborted };
    }, { moduleUrl: format === "esm" ? javascriptDataUrl(source) : undefined, fixtures: cases.map(([a, b, expected]) => [a, b, expected, [...pack({ "a.xml": a })], [...pack({ "a.xml": b })]]) });
    for (const item of result.results) {
      assert.equal(item.equal, item.expected, format);
      assert.equal(item.equal, item.oracle, format);
      assert.equal(item.strict, false, format);
      assert.equal(item.frozen, true, format);
    }
    assert.equal(result.aborted, "operation-aborted"); assert.equal(requests, 0);
    await page.close();
  }
});

test("source closure browser bundles retain unselected MCE targets and outside owners with DOM evidence", async (t) => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const parts = {
    ...alternateParts(), "outside.xml": "<outside/>",
    "_rels/outside.xml.rels": relationships([{ Id: "shared", Type: "urn:asset", Target: "shadow.bin" }]),
  };
  const good = closureBytes(parts);
  const bad = closureBytes({ ...parts, "_rels/a.xml.rels": parts["_rels/a.xml.rels"].replace('Target="shadow.bin"', 'Target="missing.bin"') });
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(); let requests = 0;
    await page.route("**/*", (route) => { requests++; return route.abort(); });
    const source = await readFile(bundlePath("pptx-parser", format), "utf8");
    if (format === "global") await page.addScriptTag({ content: source });
    const result = await page.evaluate(async ({ moduleUrl, good, bad, raw }) => {
      const parser = moduleUrl ? await import(moduleUrl) : PPTKit;
      const closure = await parser.inspectOpcClosure(new Uint8Array(good), { roots: ["/a.xml"] });
      const document = new DOMParser().parseFromString(raw, "application/xml");
      const domTargets = [...document.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship")].map((node) => node.getAttribute("Target")).sort();
      const incomplete = await parser.inspectOpcClosure(new Uint8Array(bad), { roots: ["/a.xml"] });
      const errors = [];
      for (const options of [{ roots: ["/a.xml"], closureLimits: { maxSourceRelationships: 1 } }, { roots: ["/a.xml"], closureLimits: { maxClosureExpandedBytes: 1 } }]) {
        try { await parser.inspectOpcClosure(new Uint8Array(good), options); errors.push("accepted"); }
        catch (error) { errors.push(error.code); }
      }
      return {
        names: closure.parts.map((p) => p.name), targets: closure.relationships.map((e) => e.target).sort(), domTargets,
        incoming: closure.incomingRelationships.map((e) => [e.source, e.targetPart]),
        coverage: closure.coverage, missing: incomplete.findings.map((f) => f.code),
        incomplete: incomplete.coverage.outgoingComplete, sourceCount: closure.sourceRelationshipCount,
        effectiveCount: closure.effectiveRelationshipCount, errors,
      };
    }, { moduleUrl: format === "esm" ? javascriptDataUrl(source) : undefined, good: [...good], bad: [...bad], raw: parts["_rels/a.xml.rels"] });
    assert.deepEqual(result.names, ["/_rels/a.xml.rels", "/a.xml", "/active.bin", "/shadow.bin"]);
    assert.deepEqual(result.targets, result.domTargets);
    assert.deepEqual(result.incoming, [["/outside.xml", "/shadow.bin"]]);
    assert.deepEqual(result.coverage, { outgoingComplete: true, incomingComplete: true });
    assert.deepEqual(result.missing, ["missing-source-target"]); assert.equal(result.incomplete, false);
    assert.equal(result.sourceCount, 3); assert.equal(result.effectiveCount, 2);
    assert.deepEqual(result.errors, ["package-limit-exceeded", "package-limit-exceeded"]);
    assert.equal(requests, 0); await page.close();
  }
});

test("presentation browser inspection recognizes Strict entry roots and rejects mismatches without fetching", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  let requests = 0;
  await page.route("**/*", (route) => { requests++; return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath("pptx-parser", "global"), "utf8") });
  const types = contentTypes(undefined, [{ PartName: "/root.xml", ContentType: "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml" }]);
  const make = (namespace) => pack({
    "root.xml": `<presentation xmlns="${namespace}"/>`,
    "_rels/.rels": relationships([
      { Id: "root", Type: "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument", Target: "root.xml" },
      { Id: "external", Type: "urn:custom", Target: "https://example.invalid/resource", TargetMode: "External" },
    ]),
  }, types);
  const valid = make("http://purl.oclc.org/ooxml/presentationml/main");
  const broken = make("http://schemas.openxmlformats.org/presentationml/2006/main");
  const result = await page.evaluate(async ({ valid, broken }) => {
    const identified = await PPTKit.inspectPresentationPackage(new Uint8Array(valid));
    let code;
    try { await PPTKit.inspectPresentationPackage(new Uint8Array(broken)); }
    catch (error) { code = error.code; }
    return { dialect: identified.dialect, extension: identified.suggestedExtension, main: identified.mainPart, code };
  }, { valid: [...valid], broken: [...broken] });
  assert.deepEqual(result, { dialect: "strict", extension: "potx", main: "/root.xml", code: "presentation-invalid" });
  assert.equal(requests, 0);
});

test("MCE browser projection selects fallbacks without loading namespaces or excluded external markup", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  let requests = 0;
  await page.route("**/*", (route) => { requests++; return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath("pptx-parser", "global"), "utf8") });
  const result = await page.evaluate(async () => {
    const xml = '<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="https://example.invalid/new"><mc:AlternateContent><mc:Choice Requires="x" mc:MustUnderstand="x"><x:load url="https://example.invalid/media"/></mc:Choice><mc:Fallback><n>fallback</n></mc:Fallback></mc:AlternateContent></r>';
    const value = await PPTKit.inspectMarkupCompatibility(new TextEncoder().encode(xml), { understoodNamespaces: [] });
    const errors = [];
    for (const source of [xml.replace('Requires="x"', 'Requires="missing"'), xml.replace('<r ', '<r mc:MustUnderstand="x" ')]) {
      try { await PPTKit.inspectMarkupCompatibility(new TextEncoder().encode(source), { understoodNamespaces: [] }); errors.push("accepted"); }
      catch (error) { errors.push(error.code); }
    }
    return { elements: value.output.elements, text: value.output.textCharacters, selected: value.decisions[0].selectedElementIndex, errors };
  });
  assert.deepEqual(result, { elements: 2, text: 8, selected: 5, errors: ["mce-invalid", "mce-unsupported"] });
  assert.equal(requests, 0);
});

test("browser PPTX structure follows explicit Strict slide order and rejects missing r:id targets", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  let requests = 0;
  await page.route("**/*", (route) => { requests++; return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath("pptx-parser", "global"), "utf8") });
  const fixture = withSections(structureFixture(true)), valid = encodeStructure(fixture);
  fixture.parts["deck/main.data"] = fixture.parts["deck/main.data"].replace('r:id="z"', 'r:id="absent"');
  const invalid = encodeStructure(fixture);
  const result = await page.evaluate(async ({ valid, invalid }) => {
    const report = await PPTKit.inspectPptx(new Uint8Array(valid));
    let code;
    try { await PPTKit.inspectPptx(new Uint8Array(invalid)); } catch (error) { code = error.code; }
    return { ids: report.slides.map((slide) => slide.id), hidden: report.slides.map((slide) => slide.hidden), width: report.slideSize.widthEmu, dialect: report.dialect, sections: report.sectionList.sections.map((s) => ({ id: s.id, slides: s.slides.map((slide) => slide.id) })), code };
  }, { valid: [...valid], invalid: [...invalid] });
  assert.deepEqual(result, { ids: [300, 256], hidden: [true, false], width: 9144000, dialect: "strict", sections: SECTION_IDS.map((id, i) => ({ id, slides: i === 1 ? [] : [i === 0 ? 300 : 256] })), code: "presentation-invalid" });
  assert.equal(requests, 0);
});

test("browser native section label projection matches Node without inventing a native Core collection", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  let requests = 0;
  await page.route("**/*", route => { requests++; return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath("pptx-parser", "global"), "utf8") });
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  for (const strict of [false, true]) {
    const bytes = encodeStructure(withSections(structureFixture(strict))), node = await importPptx(bytes, { projection: "structure" });
    const result = await page.evaluate(async bytes => {
      const s = await PPTKit.importPptx(new Uint8Array(bytes), { projection: "structure" });
      const report = s.report, labels = s.document.slides.map(slide => [slide.id, slide.section]);
      const plan = await s.prepareSave({ mode: "no-op" });
      if (plan.status !== "ready") throw new Error("Expected no-op plan");
      const original = Array.from((await s.save(plan.plan)).bytes), first = s.document.slides[0].id;
      s.document.moveSlide(first, 1);
      let stale;
      try { await s.save(plan.plan); } catch (error) { stale = error.code; }
      const moved = s.document.slides.map(slide => slide.section);
      s.document.moveSlide(first, 0);
      const restored = Array.from((await s.save(plan.plan)).bytes);
      s.dispose();
      return { report, labels, original, moved, stale, restored };
    }, [...bytes]);
    assert.deepEqual(result.report, node.report);
    assert.deepEqual(result.labels, node.document.slides.map(slide => [slide.id, slide.section]));
    assert.deepEqual(result.original, [...bytes]); assert.deepEqual(result.restored, [...bytes]);
    assert.deepEqual(result.moved, ["Closing", "Opening"]); assert.equal(result.stale, "save-snapshot-stale");
    node.dispose();
  }
  assert.equal(requests, 0);
});

test("opt-in native section collection agrees across Node, browser ESM and global", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath("pptx-parser", "global"), "utf8") });
  const url = javascriptDataUrl(await readFile(bundlePath("pptx-parser", "esm")));
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  for (const strict of [false, true]) {
    const bytes = encodeStructure(withSections(structureFixture(strict)));
    const node = await importPptx(bytes, { projection: "content", sectionProjection: "collection-v1" });
    const expected = { sections: node.document.sections, report: node.report };
    const results = await page.evaluate(async ({ bytes, url }) => {
      const esm = await import(url), output = [];
      for (const api of [esm, globalThis.PPTKit]) {
        const s = await api.importPptx(new Uint8Array(bytes), { projection: "content", sectionProjection: "collection-v1" });
        const plan = await s.prepareSave({ mode: "no-op" });
        if (plan.status !== "ready") throw new Error("no-op save failed");
        const saved = (await s.save(plan.plan)).bytes;
        if (saved.length !== bytes.length || saved.some((value, index) => value !== bytes[index])) throw new Error("no-op bytes changed");
        output.push({ sections: s.document.sections, report: s.report }); s.dispose();
      }
      return output;
    }, { bytes: [...bytes], url });
    assert.deepEqual(results, [expected, expected]); node.dispose();
  }
  assert.deepEqual(requests, []);
});

test("browser metadata matches Node for both dialects without fetching links or losing integer precision", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  let requests = 0;
  await page.route("**/*", (route) => { requests++; return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath("pptx-parser", "global"), "utf8") });
  const { inspectPptx } = await import("../packages/pptx-parser/dist/index.js");
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.add("core", '<d:title>中文 &amp; metadata</d:title><c:keywords xml:lang="en"><c:value>word</c:value></c:keywords>');
    f.add("extended", '<Slides>2</Slides><HyperlinkBase>https://example.invalid/base</HyperlinkBase>');
    f.add("custom", custom(2, "ui8", "18446744073709551615", 'name="__proto__" linkTarget="cached"'));
    const bytes = encodeMetadata(f);
    const result = await page.evaluate(async (bytes) => (await PPTKit.inspectPptx(new Uint8Array(bytes))).metadata, [...bytes]);
    assert.deepEqual(result, (await inspectPptx(bytes)).metadata);
    assert.equal(result.custom.properties[0].value.value, "18446744073709551615");
  }
  assert.equal(requests, 0);
});

test("standalone browser import sessions snapshot bytes, reject stale/foreign plans and never fetch resources", async (t) => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(); let requests = 0;
  await page.route("**/*", (route) => { requests++; return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath("pptx-parser", "global"), "utf8") });
  for (const strict of [false, true]) {
    const f = metadataFixture(strict); f.add("core", '<d:title>Browser import</d:title>');
    f.rootRelations.push({ Id: "external", Type: "urn:external", TargetMode: "External", Target: "https://example.invalid/never-fetch" });
    const bytes = encodeMetadata(f);
    const result = await page.evaluate(async (input) => {
      const bytes = new Uint8Array(input), importing = PPTKit.importPptx(bytes, { projection: "structure" }); bytes.fill(0);
      const a = await importing, b = await PPTKit.importPptx(new Uint8Array(input), { projection: "structure" });
      const prepared = await a.prepareSave({ mode: "no-op" }), output = await a.save(prepared.plan);
      const identical = output.bytes.length === input.length && output.bytes.every((value, i) => value === input[i]);
      let foreign, stale, disposed;
      try { await b.save(prepared.plan); } catch (e) { foreign = e.code; }
      a.document.addSlide({ id: "change" });
      try { await a.save(prepared.plan); } catch (e) { stale = e.code; }
      const blocked = await a.prepareSave({ mode: "rebuild" });
      a.dispose(); b.dispose();
      try { await a.save(prepared.plan); } catch (e) { disposed = e.code; }
      return { identical, title: a.document.metadata.title, foreign, stale, disposed, blocked: blocked.status, frozen: Object.isFrozen(a.report.inspection.slides), assurance: output.report.assurance };
    }, [...bytes]);
    assert.deepEqual(result, { identical: true, title: "Browser import", foreign: "plan-session-mismatch", stale: "save-snapshot-stale", disposed: "session-disposed", blocked: "blocked", frozen: true, assurance: "original-bytes" });
  }
  assert.equal(requests, 0);
});


test("byte assets and explicit resolvers produce matching browser and Node packages with actual image decoding", async () => {
  const { pixelPng } = await import("../tests/fixtures/presentations.mjs");
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const makeNode = () => {
    const p = createPresentation({ id: "byte-browser" });
    p.registerAsset({ id: "image", kind: "image", source: { type: "bytes", value: new Uint8Array(Buffer.from(pixelPng.split(",")[1], "base64")) } });
    p.addSlide({ id: "page" }).addElement({ id: "picture", type: "image", assetId: "image", box: { x: 0, y: 0, width: 80, height: 40 } });
    return p;
  };
  const expected = [...(await generatePptx(makeNode())).bytes];
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    for (const name of ["core", "layout", "pptx-exporter", "svg-renderer"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    const result = await page.evaluate(async (url) => {
      let fetches = 0; window.fetch = async () => { fetches++; throw new Error("unexpected fetch"); };
      const raw = Uint8Array.from(atob(url.split(",")[1]), c => c.charCodeAt(0));
      const p = PPTKit.createPresentation({ id: "byte-browser" });
      p.registerAsset({ id: "image", kind: "image", source: { type: "bytes", value: raw } }); raw.fill(0);
      p.addSlide({ id: "page" }).addElement({ id: "picture", type: "image", assetId: "image", box: { x: 0, y: 0, width: 80, height: 40 } });
      const generated = await PPTKit.generatePptx(p), preview = await PPTKit.renderPresentationToSvg(p);
      document.body.innerHTML = preview.slides[0].svg;
      const img = new Image(); img.src = document.querySelector("image").getAttribute("href"); await img.decode();
      const resolved = await PPTKit.generatePptx(p, { resolveAsset: () => ({ bytes: Uint8Array.from(atob(url.split(",")[1]), c => c.charCodeAt(0)), mimeType: "image/png" }) });
      const controller = new AbortController();
      const cancelled = PPTKit.generatePptx(p, { signal: controller.signal, resolveAsset: () => new Promise(() => {}) }); controller.abort();
      let aborted = false; try { await cancelled; } catch (e) { aborted = e.name === "AbortError"; }
      return { bytes: [...generated.bytes], resolved: [...resolved.bytes], warnings: [...generated.warnings, ...preview.warnings], size: [img.naturalWidth, img.naturalHeight], fetches, aborted };
    }, pixelPng);
    assert.deepEqual(result.bytes, expected); assert.deepEqual(result.resolved, expected); assert.deepEqual(result.warnings, []);
    assert.deepEqual(result.size, [8, 4]); assert.equal(result.fetches, 0); assert.equal(result.aborted, true);
  } finally { await browser.close(); }
});


test("browser asset lifecycle preserves references and replaces decoded preview content", async () => {
  const { pixelPng } = await import("../tests/fixtures/presentations.mjs");
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    for (const name of ["core", "layout", "pptx-exporter", "svg-renderer"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    const result = await page.evaluate(async (url) => {
      const p = PPTKit.createPresentation(), source = { type: "bytes", encoding: "base64", value: url.split(",")[1] };
      const a = p.registerAsset({ kind: "image", source });
      const deduped = p.registerAsset({ kind: "image", source }, { deduplicate: "content" });
      const slide = p.addSlide();
      for (const id of ["a", "b"]) slide.addElement({ type: "image", id, assetId: a.id, box: { x: 0, y: 0, width: 40, height: 40 } });
      let protectedUse = false; try { p.removeAsset(a.id); } catch { protectedUse = true; }
      const gif = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
      p.replaceAsset(a.id, { kind: "image", source: { type: "bytes", encoding: "base64", value: gif }, width: 1, height: 1 });
      const preview = await PPTKit.renderPresentationToSvg(p), generated = await PPTKit.generatePptx(p);
      document.body.innerHTML = preview.slides[0].svg;
      const image = new Image(); image.src = document.querySelector("image").getAttribute("href"); await image.decode();
      const references = p.getAssetReferences(a.id).map(ref => ref.elementId);
      for (const element of slide.elements) slide.removeElement(element.id);
      const kept = p.removeUnusedAssets({ retainAssetIds: [a.id] }).length;
      const removed = p.removeUnusedAssets().map(asset => asset.id);
      return { deduplicated: a === deduped, protectedUse, references, size: [image.naturalWidth, image.naturalHeight], warnings: [...preview.warnings, ...generated.warnings], kept, removed: removed.length, assets: p.assets.length, oldSource: a.source.value === source.value };
    }, pixelPng);
    assert.deepEqual(result, { deduplicated: true, protectedUse: true, references: ["a", "b"], size: [1, 1], warnings: [], kept: 0, removed: 1, assets: 0, oldSource: true });
  } finally { await browser.close(); }
});


test("browser document copies remap connector identities and self links before rendering/export", async () => {
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    for (const name of ["core", "layout", "pptx-exporter", "svg-renderer"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    const result = await page.evaluate(async () => {
      const p = PPTKit.createPresentation(), s = p.addSlide({ id: "source" });
      s.addElement({ type: "connector", id: "edge", start: { elementId: "target", anchor: "right" }, end: { x: 100, y: 50 } });
      s.addElement({ type: "shape", shape: "rect", id: "target", box: { x: 10, y: 10, width: 20, height: 20 }, action: { type: "slide", slideId: "source" } });
      const copy = p.duplicateSlide(s.id); p.removeSlide(s.id);
      const layout = PPTKit.resolveLayout(p), svg = await PPTKit.renderPresentationToSvg(p), output = await PPTKit.generatePptx(p);
      return { errors: PPTKit.validatePresentation(p), linked: copy.elements[0].start.elementId === copy.elements[1].id, selfLink: copy.elements[1].action.slideId === copy.id, start: layout.slides[0].elements[0].start, warnings: output.warnings, visible: svg.slides[0].svg.includes("30,20 100,50") };
    });
    assert.deepEqual(result, { errors: [], linked: true, selfLink: true, start: { x: 30, y: 20 }, warnings: [], visible: true });
  } finally { await browser.close(); }
});


test("native connector mode is byte-equivalent in browser and Node and retains layout endpoint identities", async () => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const make = create => {
    const p = create({ id: "native-browser" }), s = p.addSlide({ id: "slide" });
    s.addElement({ type: "connector", id: "edge", start: { elementId: "a", anchor: "right" }, end: { elementId: "b", anchor: "left" } });
    s.addElement({ type: "connector", id: "elbow", start: { x: 80, y: 120 }, end: { x: 20, y: 60 }, route: [{ x: 20, y: 120 }] });
    s.addElement({ type: "shape", id: "a", shape: "rect", box: { x: 0, y: 0, width: 40, height: 20 } });
    s.addElement({ type: "shape", id: "b", shape: "ellipse", box: { x: 100, y: 0, width: 40, height: 20 } }); return p;
  };
  const expected = [...(await generatePptx(make(createPresentation), { connectorMode: "native" })).bytes];
  assert.match(readGeneratedZip(Uint8Array.from(expected)).get("ppt/slides/slide1.xml").toString(), /prst="bentConnector2"/);
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    await page.addScriptTag({ content: `window.makeNativeConnectorDeck = ${make.toString()};` });
    const result = await page.evaluate(async () => {
      const p = window.makeNativeConnectorDeck(PPTKit.createPresentation), layout = PPTKit.resolveLayout(p), generated = await PPTKit.generatePptx(p, { connectorMode: "native" });
      return { bytes: [...generated.bytes], warnings: generated.warnings, connections: layout.slides[0].elements[0].connections };
    });
    assert.deepEqual(result.bytes, expected); assert.deepEqual(result.warnings, []);
    assert.deepEqual(result.connections, { start: { elementId: "a", anchor: "right" }, end: { elementId: "b", anchor: "left" } });
  } finally { await browser.close(); }
});

test("content import and reports match Node in a real browser and retain the exact no-op source", async () => {
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const { normalizePresentation } = await import("../packages/core/dist/index.js");
  const input = contentBytes(), node = await importPptx(input, { projection: "content" });
  const expected = { document: normalizePresentation(node.document), report: node.report };
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    for (const name of ["core", "pptx-parser", "layout", "svg-renderer", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    const result = await page.evaluate(async input => {
      const session = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
      const document = PPTKit.normalizePresentation(session.document), report = session.report;
      const prepared = await session.prepareSave({ mode: "no-op" }), saved = await session.save(prepared.plan);
      const svg = await PPTKit.renderPresentationToSvg(session.document), fresh = await PPTKit.generatePptx(session.document);
      const importedAgain = await PPTKit.importPptx(fresh.bytes, { projection: "content" });
      const text = importedAgain.document.slides[0].elements[0].content[0].runs[0].text;
      const blocked = await session.prepareSave({ mode: "edit" });
      session.dispose(); importedAgain.dispose();
      return { document, report, saved: [...saved.bytes], text, hasText: svg.slides[0].svg.includes("A &amp; B"), blocked: blocked.status };
    }, [...input]);
    assert.deepEqual(result.document, expected.document); assert.deepEqual(result.report, expected.report);
    assert.deepEqual(result.saved, [...input]); assert.equal(result.text, " A & B "); assert.equal(result.hasText, true); assert.equal(result.blocked, "blocked");
  } finally { node.dispose(); await browser.close(); }
});

test("unsupported image MIME bypasses projection budget identically in Chromium and Node", async t => {
  const { importPptx } = await import('../packages/pptx-parser/dist/index.js');
  const fixture = imageFixture(picture({ id: 2, embed: 'vector' }) + picture({ id: 3, embed: 'image' }));
  fixture.parts['media/vector.svg'] = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg">' + ' '.repeat(png.length * 3) + '</svg>');
  fixture.mediaTypes['media/vector.svg'] = 'image/svg+xml';
  fixture.parts['deck/_rels/z.xml.rels'] = relationships([fixture.rel('layout', 'slideLayout', 'layout.xml'), fixture.rel('image', 'image', '../media/raster.bin'), fixture.rel('vector', 'image', '../media/vector.svg')]);
  const bytes = encodeImages(fixture), options = { projection: 'content', sessionLimits: { maxImportedAssetBytes: png.length, maxTotalImportedAssetBytes: png.length } };
  const node = await importPptx(bytes, options), report = node.report; node.dispose();
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (const name of ['core', 'pptx-parser']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const actual = await page.evaluate(async ({ bytes, options }) => {
    const session = await PPTKit.importPptx(Uint8Array.from(bytes), options);
    const result = { report: session.report, count: session.document.assets.length, saved: [...(await session.save((await session.prepareSave({ mode: 'no-op' })).plan)).bytes] };
    session.dispose(); return result;
  }, { bytes: [...bytes], options });
  assert.deepEqual(actual.report, report); assert.equal(actual.count, 1); assert.deepEqual(actual.saved, [...bytes]);
});

test("text save retains shared and opaque pictures identically in Chromium and Node", async t => {
  const { importPptx } = await import('../packages/pptx-parser/dist/index.js');
  const extension='<a:extLst><a:ext uri="urn:test"><v:svgBlip xmlns:v="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="svg"/></a:ext></a:extLst>';
  const fixture=imageFixture(shape({id:'2',txBox:true,text:body()})+picture({id:3,effects:extension})+picture({id:5})+picture({id:4,embed:'legacy'}));
  fixture.parts['media/vector.svg']=Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>');fixture.mediaTypes['media/vector.svg']='image/svg+xml';
  fixture.parts['media/legacy.wmf']=await readFile(path.join(root,'tests/fixtures/metafile-libreoffice-asymmetric.wmf'));fixture.mediaTypes['media/legacy.wmf']='image/wmf';
  fixture.parts['deck/_rels/z.xml.rels']=relationships([fixture.rel('layout','slideLayout','layout.xml'),fixture.rel('image','image','../media/raster.bin'),fixture.rel('svg','image','../media/vector.svg'),fixture.rel('legacy','image','../media/legacy.wmf')]);
  const bytes=encodeImages(fixture),options={projection:'content',editSaveProfile:'text-runs-v1'},node=await importPptx(bytes,options);
  const slide=node.document.slides[0],element=structuredClone(slide.elements[0]);element.content[0].runs[0].text='Changed';slide.replaceElement(element.id,element);
  const plan=await node.prepareSave({mode:'edit'});assert.equal(plan.status,'ready');const expected=await node.save(plan.plan);node.dispose();
  const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();let requests=0;await page.route('**/*',r=>{requests++;return r.abort()});
  for(const name of ['core','pptx-parser'])await page.addScriptTag({content:await readFile(bundlePath(name,'global'),'utf8')});
  const actual=await page.evaluate(async({bytes,options})=>{const s=await PPTKit.importPptx(Uint8Array.from(bytes),options);try{const slide=s.document.slides[0],element=structuredClone(slide.elements[0]);element.content[0].runs[0].text='Changed';slide.replaceElement(element.id,element);const plan=await s.prepareSave({mode:'edit'});const output=await s.save(plan.plan);return {bytes:[...output.bytes],report:output.report}}finally{s.dispose()}},{bytes:[...bytes],options});
  assert.deepEqual(actual.bytes,[...expected.bytes]);assert.deepEqual(actual.report,expected.report);assert.equal(requests,0);
});

test("SVG extension raster-only import warning and no-op match Chromium and Node", async t => {
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const extension = '<a:extLst><a:ext uri="urn:test"><v:svgBlip xmlns:v="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:link="svg"/></a:ext></a:extLst>';
  const fixture = imageFixture(picture({ effects: extension }));
  fixture.parts["deck/_rels/z.xml.rels"] = relationships([fixture.rel("layout", "slideLayout", "layout.xml"), fixture.rel("image", "image", "../media/raster.bin"), { ...fixture.rel("svg", "image", "https://example.invalid/vector.svg"), TargetMode: "External" }]);
  const input = encodeImages(fixture), node = await importPptx(input, { projection: "content" });
  const expected = node.report; node.dispose();
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(); let requests = 0;
  await page.route("**/*", route => { requests++; return route.abort(); });
  for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const result = await page.evaluate(async input => {
    const session = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
    const report = session.report, saved = await session.save((await session.prepareSave({ mode: "no-op" })).plan);
    session.dispose(); return { report, saved: [...saved.bytes] };
  }, [...input]);
  assert.deepEqual(result.report, expected);
  assert.ok(result.report.findings.some(f => /An SVG image extension is present/.test(f.reason)));
  assert.deepEqual(result.saved, [...input]); assert.equal(requests, 0);
});

test("imported raster images render identically in Chromium, share bytes with Node and never load external-only pictures", async () => {
  const { createPresentation, normalizePresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const input = (await generatePptx(createImageImportFixture(createPresentation, png))).bytes;
  const node = await importPptx(input, { projection: "content" }), expected = { document: normalizePresentation(node.document), report: node.report };
  const external = imageFixture(picture({ embed: null, link: "remote" }));
  external.parts["deck/_rels/z.xml.rels"] = relationships([external.rel("layout", "slideLayout", "layout.xml"), { ...external.rel("remote", "image", "https://example.invalid/unrequested.png"), TargetMode: "External" }]);
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    for (const name of ["core", "layout", "pptx-parser", "svg-renderer", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    await page.addScriptTag({ content: `window.makeImageImportFixture = ${createImageImportFixture.toString()};` });
    const result = await page.evaluate(async ({ input, png, external }) => {
      const s = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
      const reference = window.makeImageImportFixture(PPTKit.createPresentation, Uint8Array.from(png));
      const a = await PPTKit.renderPresentationToSvg(reference), b = await PPTKit.renderPresentationToSvg(s.document);
      const pixels = async svg => {
        const image = new Image(); image.src = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`; await image.decode();
        const canvas = globalThis.document.createElement("canvas"); canvas.width = 400; canvas.height = 240;
        const ctx = canvas.getContext("2d"); ctx.drawImage(image, 0, 0, 400, 240);
        return [...ctx.getImageData(0, 0, 400, 240).data];
      };
      const firstPixels = await pixels(a.slides[0].svg), projectedPixels = await pixels(b.slides[0].svg);
      const decoded = new Image(); decoded.src = `data:image/png;base64,${s.document.assets[0].source.value}`; await decoded.decode();
      const fresh = await PPTKit.generatePptx(s.document), saved = await s.save((await s.prepareSave({ mode: "no-op" })).plan);
      const e = await PPTKit.importPptx(Uint8Array.from(external), { projection: "content" });
      const externalPreview = await PPTKit.renderPresentationToSvg(e.document);
      const externalExport = await PPTKit.generatePptx(e.document);
      const result = { document: PPTKit.normalizePresentation(s.document), report: s.report, exactPixels: firstPixels.every((v, i) => v === projectedPixels[i]), renderedNonWhite: projectedPixels.some((v, i) => i % 4 !== 3 && v !== 255), dimensions: [decoded.naturalWidth, decoded.naturalHeight], fresh: [...fresh.bytes], saved: [...saved.bytes], externalAssets: e.document.assets.length, externalElements: e.document.slides[0].elements.length, externalUrlInPreview: externalPreview.slides.some(slide => slide.svg.includes("example.invalid")), externalWarnings: externalExport.warnings };
      s.dispose(); e.dispose(); return result;
    }, { input: [...input], png: [...png], external: [...encodeImages(external)] });
    assert.deepEqual(result.document, expected.document); assert.deepEqual(result.report, expected.report);
    assert.equal(result.exactPixels, true); assert.equal(result.renderedNonWhite, true); assert.deepEqual(result.dimensions, [8, 4]);
    assert.deepEqual(result.saved, [...input]); assert.deepEqual(result.fresh, [...(await generatePptx(node.document)).bytes]);
    assert.equal(result.externalAssets, 0); assert.equal(result.externalElements, 0);
    assert.equal(result.externalUrlInPreview, false); assert.deepEqual(result.externalWarnings, []); assert.deepEqual(requests, []);
  } finally { node.dispose(); await browser.close(); }
});

test("browser connector/action import matches Node and regenerates moved attached targets without activating links", async () => {
  const { createPresentation, normalizePresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const input = (await generatePptx(createReferenceImportFixture(createPresentation), { connectorMode: "native" })).bytes;
  const node = await importPptx(input, { projection: "content" }), expected = { document: normalizePresentation(node.document), report: node.report, bytes: [...(await generatePptx(node.document, { connectorMode: "native" })).bytes] };
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    for (const name of ["core", "layout", "pptx-parser", "svg-renderer", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    const result = await page.evaluate(async input => {
      const s = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
      const document = PPTKit.normalizePresentation(s.document), report = s.report, prepared = await s.prepareSave({ mode: "no-op" });
      const noOp = await s.save(prepared.plan), fresh = await PPTKit.generatePptx(s.document, { connectorMode: "native" });
      const rendered = await PPTKit.renderPresentationToSvg(s.document);
      globalThis.document.body.innerHTML = rendered.slides[0].svg;
      const links = [...globalThis.document.querySelectorAll("a[href]")].map(a => a.getAttribute("href"));
      const slideTargets = [...globalThis.document.querySelectorAll("[data-pptkit-slide-target]")].map(e => e.getAttribute("data-pptkit-slide-target"));
      const slide = s.document.slides[0], target = slide.removeElement(slide.elements[1].id);
      slide.insertElement(1, { ...target, box: { x: 100, y: 100, width: 40, height: 40 } });
      let stale;
      try { await s.save(prepared.plan); } catch (e) { stale = e.code; }
      const updated = await PPTKit.generatePptx(s.document, { connectorMode: "native" }), again = await PPTKit.importPptx(updated.bytes, { projection: "content" });
      const moved = PPTKit.resolveLayout(again.document).slides[0].elements[0].start;
      const end = PPTKit.resolveLayout(again.document).slides[0].elements[0].end;
      const blocked = await s.prepareSave({ mode: "edit" });
      s.dispose(); again.dispose();
      return { document, report, fresh: [...fresh.bytes], noOp: [...noOp.bytes], links, slideTargets, stale, moved, end, blocked: blocked.status, warnings: rendered.warnings.map(w => w.code) };
    }, [...input]);
    assert.deepEqual(result.document, expected.document); assert.deepEqual(result.report, expected.report);
    assert.deepEqual(result.fresh, expected.bytes); assert.deepEqual(result.noOp, [...input]);
    assert.ok(result.links.includes("https://example.invalid/a?x=1&y=%E4%B8%AD"));
    assert.ok(result.links.includes("mailto:user@example.invalid?subject=Hello%20world"));
    assert.ok(result.slideTargets.includes(node.document.slides[0].id)); assert.ok(result.slideTargets.includes(node.document.slides[1].id));
    assert.equal(result.stale, "save-snapshot-stale"); assert.deepEqual(result.moved, { x: 140, y: 120 });
    assert.deepEqual(result.end, { x: 230, y: 90 }); assert.equal(result.blocked, "blocked");
    assert.ok(result.warnings.includes("slide-action-not-linked")); assert.deepEqual(requests, []);
  } finally { node.dispose(); await browser.close(); }
});

test("browser speaker notes import, editing and fresh output match Node without exposing notes in slide SVG or loading links", async () => {
  const { createPresentation, normalizePresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const input = (await generatePptx(createNotesImportFixture(createPresentation))).bytes;
  const node = await importPptx(input, { projection: "content" });
  const expected = { document: normalizePresentation(node.document), report: node.report, fresh: [...(await generatePptx(node.document)).bytes] };
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    for (const name of ["core", "layout", "pptx-parser", "svg-renderer", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    const result = await page.evaluate(async input => {
      const s = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
      const document = PPTKit.normalizePresentation(s.document), report = s.report;
      const prepared = await s.prepareSave({ mode: "no-op" }), saved = await s.save(prepared.plan);
      const fresh = await PPTKit.generatePptx(s.document), rendered = await PPTKit.renderPresentationToSvg(s.document);
      globalThis.document.body.innerHTML = rendered.slides[0].svg;
      const exposed = globalThis.document.body.textContent.includes("Speaker α");
      const slide = s.document.slides[0], original = slide.notes;
      slide.setNotes([{ runs: [{ text: "Browser edit" }] }]);
      let stale; try { await s.save(prepared.plan); } catch (e) { stale = e.code; }
      const updated = await PPTKit.generatePptx(s.document), again = await PPTKit.importPptx(updated.bytes, { projection: "content" });
      const edited = again.document.slides[0].notes[0].runs[0].text;
      slide.setNotes(original);
      const restored = await s.save(prepared.plan);
      s.dispose(); again.dispose();
      return { document, report, fresh: [...fresh.bytes], saved: [...saved.bytes], exposed, stale, edited, restored: [...restored.bytes] };
    }, [...input]);
    assert.deepEqual(result.document, expected.document); assert.deepEqual(result.report, expected.report);
    assert.deepEqual(result.fresh, expected.fresh); assert.deepEqual(result.saved, [...input]); assert.deepEqual(result.restored, [...input]);
    assert.equal(result.exposed, false); assert.equal(result.stale, "save-snapshot-stale"); assert.equal(result.edited, "Browser edit");
    assert.deepEqual(requests, []);
  } finally { node.dispose(); await browser.close(); }
});

test("merged table import matches Node bytes/reports and actual browser cell pixels without loading links", async () => {
  const { createPresentation, normalizePresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const input = (await generatePptx(createTableImportFixture(createPresentation))).bytes;
  const node = await importPptx(input, { projection: "content" });
  const expected = { document: normalizePresentation(node.document), report: node.report, fresh: [...(await generatePptx(node.document)).bytes] };
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    for (const name of ["core", "layout", "pptx-parser", "svg-renderer", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    await page.addScriptTag({ content: `window.makeTableImportFixture = ${createTableImportFixture.toString()};` });
    const result = await page.evaluate(async input => {
      const s = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
      const document = PPTKit.normalizePresentation(s.document), report = s.report;
      const prepared = await s.prepareSave({ mode: "no-op" }), saved = await s.save(prepared.plan);
      const fresh = await PPTKit.generatePptx(s.document);
      const reference = await PPTKit.renderPresentationToSvg(window.makeTableImportFixture(PPTKit.createPresentation));
      const projected = await PPTKit.renderPresentationToSvg(s.document);
      window.__tableSvgs = [reference.slides[0].svg, projected.slides[0].svg];
      globalThis.document.body.style.margin = "0";
      globalThis.document.body.innerHTML = `<div id="table-surface" style="width:640px;height:480px">${projected.slides[0].svg}</div>`;
      await globalThis.document.fonts.ready;
      const cells = [...globalThis.document.querySelectorAll("td, [data-pptkit-cell]")].map(cell => ({ text: cell.textContent, rows: cell.rowSpan ?? Number(cell.getAttribute("aria-rowspan")), columns: cell.colSpan ?? Number(cell.getAttribute("aria-colspan")), width: cell.getBoundingClientRect().width, height: cell.getBoundingClientRect().height }));
      const slide = s.document.slides[0], table = slide.removeElement(slide.elements[0].id), changed = structuredClone(table);
      (changed.type === "group" ? changed.children[0] : changed).rows[2].cells[0].content = "Browser table edit"; slide.insertElement(0, changed);
      let stale; try { await s.save(prepared.plan); } catch (e) { stale = e.code; }
      const updated = await PPTKit.generatePptx(s.document), again = await PPTKit.importPptx(updated.bytes, { projection: "content" });
      const root = again.document.slides[0].elements[0];
      const edited = (root.type === "group" ? root.children[0] : root).rows[2].cells[0].content[0].runs[0].text;
      slide.removeElement(table.id); slide.insertElement(0, table);
      const restored = await s.save(prepared.plan);
      s.dispose(); again.dispose();
      return { document, report, saved: [...saved.bytes], fresh: [...fresh.bytes], cells, stale, edited, restored: [...restored.bytes] };
    }, [...input]);
    assert.deepEqual(result.document, expected.document); assert.deepEqual(result.report, expected.report);
    assert.deepEqual(result.saved, [...input]); assert.deepEqual(result.restored, [...input]); assert.deepEqual(result.fresh, expected.fresh);
    assert.equal(result.cells.length, 7);
    assert.ok(result.cells[0].text.includes("Merged α😀 & <literal>")); assert.equal(result.cells[0].rows, 2); assert.equal(result.cells[0].columns, 2);
    assert.ok(result.cells.every(cell => cell.width > 0 && cell.height > 0));
    assert.equal(result.stale, "save-snapshot-stale"); assert.equal(result.edited, "Browser table edit");
    await page.evaluate(() => { document.querySelector("#table-surface").innerHTML = window.__tableSvgs[0]; });
    const originalPixels = await page.locator("#table-surface").screenshot({ animations: "disabled" });
    await page.evaluate(() => { document.querySelector("#table-surface").innerHTML = window.__tableSvgs[1]; });
    const projectedPixels = await page.locator("#table-surface").screenshot({ animations: "disabled" });
    const difference = await page.evaluate(async images => {
      const decode = async data => {
        const image = new Image(); image.src = `data:image/png;base64,${data}`; await image.decode();
        const canvas = document.createElement("canvas"); canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
        const context = canvas.getContext("2d"); context.drawImage(image, 0, 0);
        return context.getImageData(0, 0, canvas.width, canvas.height);
      };
      const [a, b] = await Promise.all(images.map(decode));
      if (a.width !== b.width || a.height !== b.height) return { different: Infinity, max: Infinity };
      let different = 0, max = 0;
      for (let i = 0; i < a.data.length; i += 4) {
        const delta = Math.max(...[0, 1, 2, 3].map(channel => Math.abs(a.data[i + channel] - b.data[i + channel])));
        if (delta) different++;
        max = Math.max(max, delta);
      }
      return { different, max };
    }, [originalPixels.toString("base64"), projectedPixels.toString("base64")]);
    assert.ok(difference.different <= 10 && difference.max <= 1, `reference/imported table pixels differ beyond raster rounding: ${JSON.stringify(difference)}`);
    assert.deepEqual(requests, []);
  } finally { node.dispose(); await browser.close(); }
});


test("chart content data/workbook import, edits and reports agree in Chromium and Node without network requests", async () => {
  const { createPresentation, normalizePresentation } = await import("../packages/core/dist/index.js");
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const input = (await generatePptx(createChartImportFixture(createPresentation))).bytes;
  const node = await importPptx(input, { projection: "content" });
  const expected = { document: normalizePresentation(node.document), report: node.report, fresh: [...(await generatePptx(node.document)).bytes] };
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    for (const name of ["core", "layout", "pptx-parser", "svg-renderer", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    const result = await page.evaluate(async input => {
      const s = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
      const document = PPTKit.normalizePresentation(s.document), report = s.report;
      const prepared = await s.prepareSave({ mode: "no-op" }), saved = await s.save(prepared.plan), fresh = await PPTKit.generatePptx(s.document);
      const svg = await PPTKit.renderPresentationToSvg(s.document);
      const slide = s.document.slides[0], previous = slide.removeElement(slide.elements[0].id), changed = structuredClone(previous);
      changed.series = [{ name: "Browser chart edit", values: [6, 7, 8] }]; slide.insertElement(0, changed);
      let stale; try { await s.save(prepared.plan); } catch (e) { stale = e.code; }
      const updated = await PPTKit.generatePptx(s.document), again = await PPTKit.importPptx(updated.bytes, { projection: "content" });
      const edited = again.document.slides[0].elements[0].series.map(s => [s.name, s.values]);
      const blocked = (await s.prepareSave({ mode: "edit" })).status;
      s.dispose(); again.dispose(); return { document, report, fresh: [...fresh.bytes], saved: [...saved.bytes], svg: svg.slides[0].svg, edited, stale, blocked };
    }, [...input]);
    assert.deepEqual(result.document, expected.document); assert.deepEqual(result.report, expected.report);
    assert.deepEqual(result.saved, [...input]); assert.deepEqual(result.fresh, expected.fresh);
    assert.match(result.svg, /Chart α😀/); assert.deepEqual(result.edited, [["Browser chart edit", [6, 7, 8]]]);
    assert.equal(result.stale, "save-snapshot-stale"); assert.equal(result.blocked, "blocked"); assert.deepEqual(requests, []);
  } finally { node.dispose(); await browser.close(); }
});

test("placeholder matching/provenance and fresh geometry agree in Chromium and Node without inherited text duplication", async () => {
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const { normalizePresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const input = placeholderBytes(true), node = await importPptx(input, { projection: "content" });
  const expected = { document: normalizePresentation(node.document), report: node.report, fresh: [...(await generatePptx(node.document)).bytes] };
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    for (const name of ["core", "layout", "pptx-parser", "svg-renderer", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    const result = await page.evaluate(async input => {
      const s = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content", sessionLimits: { maxPlaceholderEntries: 4 } });
      const document = PPTKit.normalizePresentation(s.document), report = s.report, fresh = await PPTKit.generatePptx(s.document);
      const preview = await PPTKit.renderPresentationToSvg(s.document);
      globalThis.document.body.innerHTML = preview.slides[0].svg;
      const text = globalThis.document.body.textContent;
      const prepared = await s.prepareSave({ mode: "no-op" }), saved = await s.save(prepared.plan);
      let budget;
      try { await PPTKit.importPptx(Uint8Array.from(input), { projection: "content", sessionLimits: { maxPlaceholderEntries: 3 } }); }
      catch (e) { budget = e.code; }
      s.dispose(); return { document, report, fresh: [...fresh.bytes], saved: [...saved.bytes], text, budget };
    }, [...input]);
    assert.deepEqual(result.document, expected.document); assert.deepEqual(result.report, expected.report);
    assert.deepEqual(result.fresh, expected.fresh); assert.deepEqual(result.saved, [...input]);
    assert.match(result.text, /A & B/); assert.doesNotMatch(result.text, /MASTER PROMPT/);
    assert.equal(result.budget, "import-limit-exceeded"); assert.deepEqual(requests, []);
  } finally { node.dispose(); await browser.close(); }
});

test("partial layout theme overrides preserve independent color/font resolution in Chromium and Node", async () => {
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const { normalizePresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const fixture = setLayoutThemeFont(setTextTheme(textStyleFixture(true), { accent1: "3579BD", major: "Base Browser Font" }), "Override Browser Font");
  fixture.parts["deck/layout.xml"] = fixture.parts["deck/layout.xml"].replace(layoutList,
    '<a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr>');
  const input = encodeStructure(fixture), node = await importPptx(input, { projection: "content" });
  const expected = { document: normalizePresentation(node.document), report: node.report, fresh: [...(await generatePptx(node.document)).bytes] };
  fixture.types["deck/font-override.xml"] = "theme";
  const invalid = encodeStructure(fixture), browser = await chromium.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    for (const name of ["core", "layout", "pptx-parser", "svg-renderer", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
    const result = await page.evaluate(async ({ input, invalid }) => {
      const session = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
      try {
        const document = PPTKit.normalizePresentation(session.document), report = session.report;
        const plan = await session.prepareSave({ mode: "no-op" }), saved = await session.save(plan.plan);
        const fresh = await PPTKit.generatePptx(session.document);
        const reopened = await PPTKit.importPptx(fresh.bytes, { projection: "content" });
        const reopenedDocument = PPTKit.normalizePresentation(reopened.document);
        reopened.dispose();
        let rejected;
        try { await PPTKit.importPptx(Uint8Array.from(invalid), { projection: "content" }); }
        catch (error) { rejected = error.code; }
        return { document, report, saved: [...saved.bytes], fresh: [...fresh.bytes], reopenedDocument, rejected };
      } finally { session.dispose(); }
    }, { input: [...input], invalid: [...invalid] });
    assert.deepEqual(result.document, expected.document);
    assert.deepEqual(result.report, expected.report);
    assert.deepEqual(result.fresh, expected.fresh);
    assert.deepEqual(result.saved, [...input]);
    for (const document of [result.document, result.reopenedDocument]) for (const slide of document.slides) {
      assert.equal(slide.elements[0].text.content[0].runs[0].style.color, "3579BD");
      assert.equal(slide.elements[0].text.content[0].runs[0].style.fontFamily, "Override Browser Font");
    }
    assert.ok(result.report.textStyleInheritance.properties.some(property => property.property === "run.fontFamily"
      && property.resolvedFrom?.some(source => source.part === "/deck/font-override.xml")));
    assert.equal(result.rejected, "import-invalid-package");
    assert.deepEqual(requests, []);
  } finally { node.dispose(); await browser.close(); }
});

test("private XML character edits retain exact encodings in Chromium and agree with independent DOM parsing", async () => {
  // A test-only bundle exercises the private prerequisite without exporting it
  // from the published Parser API.
  const { rollup } = await import("rollup");
  const { default: nodeResolve } = await import("@rollup/plugin-node-resolve");
  const { default: commonjs } = await import("@rollup/plugin-commonjs");
  const bundle = await rollup({
    input: path.join(root, "packages/pptx-parser/dist/preservation/xml-text-source.js"),
    plugins: [nodeResolve({ browser: true }), commonjs()],
  });
  let source;
  try { source = (await bundle.generate({ format: "iife", name: "XmlSourceTest" })).output[0].code; }
  finally { await bundle.close(); }
  assert.doesNotMatch(source, /(?:^|["'])node:|\brequire\s*\(|\bprocess(?:\.|\[)/);
  const inputs = ["utf-8", "utf-16le", "utf-16be"].map(encoding => {
    const before = `<?xml version="1.0" encoding="${encoding === "utf-8" ? "UTF-8" : "UTF-16"}"?>\r\n<a:r xmlns:a="urn:a" xmlns:f="urn:future"><!--keep <a:t/>--><a:t><![CDATA[old]]></a:t><f:effect ref="a:t">keep</f:effect><a:t /></a:r>`;
    const after = before.replace("<![CDATA[old]]>", "browser&lt;&amp;&#xD;🙂").replace("<a:t />", "<a:t >empty</a:t>");
    const encode = text => encoding === "utf-8" ? Buffer.concat([Buffer.from([239, 187, 191]), Buffer.from(text)]) : Buffer.concat([
      Buffer.from(encoding === "utf-16le" ? [255, 254] : [254, 255]),
      encoding === "utf-16le" ? Buffer.from(text, "utf16le") : Buffer.from(text, "utf16le").swap16(),
    ]);
    return { encoding, input: [...encode(before)], expected: [...encode(after)] };
  });
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    await page.addScriptTag({ content: source });
    const result = await page.evaluate(async inputs => {
      const outputs = [];
      for (const { encoding, input } of inputs) {
        const captured = await XmlSourceTest.captureXmlTextSource(Uint8Array.from(input));
        const targets = captured.elements.filter(e => e.namespaceUri === "urn:a" && e.local === "t");
        const result = await captured.replaceText(targets.map((e, i) => ({
          elementIndex: e.elementIndex, namespaceUri: e.namespaceUri, local: e.local,
          expectedText: e.text, text: i === 0 ? "browser<&\r🙂" : "empty",
        })));
        const xml = new DOMParser().parseFromString(new TextDecoder(encoding).decode(result.bytes), "application/xml");
        const texts = [...xml.getElementsByTagNameNS("urn:a", "t")].map(e => e.textContent);
        const future = xml.getElementsByTagNameNS("urn:future", "effect")[0];
        let budget;
        try { await captured.replaceText([], { limits: { maxOutputBytes: 1 } }); }
        catch (error) { budget = error.code; }
        captured.dispose();
        let disposed;
        try { await captured.replaceText([]); } catch (error) { disposed = error.code; }
        outputs.push({
          bytes: [...result.bytes], texts, future: [future.getAttribute("ref"), future.textContent],
          errors: xml.getElementsByTagName("parsererror").length, budget, disposed,
        });
      }
      return outputs;
    }, inputs.map(({ encoding, input }) => ({ encoding, input })));
    for (let i = 0; i < inputs.length; i++) {
      assert.deepEqual(result[i].bytes, inputs[i].expected);
      assert.deepEqual(result[i].texts, ["browser<&\r🙂", "empty"]);
      assert.deepEqual(result[i].future, ["a:t", "keep"]);
      assert.equal(result[i].errors, 0);
      assert.equal(result[i].budget, "xml-edit-limit-exceeded");
      assert.equal(result[i].disposed, "xml-source-disposed");
    }
    assert.deepEqual(requests, []);
  } finally { await browser.close(); }
});

test("public text-runs-v1 save plans agree in Node, browser ESM/global and independent source DOM identities", async () => {
  const { createPresentation, normalizePresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const { readGeneratedZip } = await import("../tests/support/read-generated-zip.mjs");
  const input = (await generatePptx(createReferenceImportFixture(createPresentation), { connectorMode: "native" })).bytes;
  const sourceXml = readGeneratedZip(input).get("ppt/slides/slide1.xml").toString();
  const node = await importPptx(input, { projection: "content", editSaveProfile: "text-runs-v1" });
  const slide = node.document.slides[0], index = slide.elements.findIndex(e => e.type === "text"), text = structuredClone(slide.removeElement(slide.elements[index].id));
  text.content[0].runs[0].text = "Preserved browser <&🙂"; slide.insertElement(index, text);
  const prepared = await node.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
  const saved = await node.save(prepared.plan), expected = { bytes: [...saved.bytes], report: saved.report, document: normalizePresentation(node.document), bindings: node.report.textRuns };
  node.dispose();
  const browser = await chromium.launch();
  try {
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = [];
      await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(bundlePath(name, format), "utf8") });
      const esm = format === "esm" ? {
        core: javascriptDataUrl(await readFile(bundlePath("core", format), "utf8")),
        parser: javascriptDataUrl(await readFile(bundlePath("pptx-parser", format), "utf8")),
      } : undefined;
      const result = await page.evaluate(async ({ input, sourceXml, esm }) => {
        const api = esm ? await import(esm.parser) : PPTKit, core = esm ? await import(esm.core) : PPTKit;
        const session = await api.importPptx(Uint8Array.from(input), { projection: "content", editSaveProfile: "text-runs-v1" });
        try {
          const dom = new DOMParser().parseFromString(sourceXml, "application/xml"), elements = [...dom.getElementsByTagName("*")];
          const sourceMatches = session.report.textRuns.every(binding => {
            const element = elements[binding.sourceElementIndex - 1];
            return element?.localName === "t" && element.namespaceURI.endsWith("/main");
          });
          const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "text");
          const text = structuredClone(slide.removeElement(slide.elements[index].id)); text.content[0].runs[0].text = "Preserved browser <&🙂"; slide.insertElement(index, text);
          const plan = await session.prepareSave({ mode: "edit" });
          if (plan.status !== "ready") throw new Error(JSON.stringify(plan.report));
          const saved = await session.save(plan.plan), document = core.normalizePresentation(session.document), bindings = session.report.textRuns;
          const changed = structuredClone(slide.removeElement(text.id)); changed.content[0].runs[0].text = "stale"; slide.insertElement(index, changed);
          let stale; try { await session.save(plan.plan); } catch (error) { stale = error.code; }
          const controller = new AbortController(), listeners = new Set();
          let notifySubscription;
          const subscribed = new Promise(resolve => { notifySubscription = resolve; });
          const add = controller.signal.addEventListener.bind(controller.signal), remove = controller.signal.removeEventListener.bind(controller.signal);
          controller.signal.addEventListener = (type, listener, options) => {
            add(type, listener, options);
            if (type === "abort") { listeners.add(listener); notifySubscription(); }
          };
          controller.signal.removeEventListener = (type, listener, options) => {
            remove(type, listener, options);
            if (type === "abort") listeners.delete(listener);
          };
          const pending = session.prepareSave({ mode: "edit", signal: controller.signal })
            .then(() => "unexpected-ready", error => error.code);
          await subscribed;
          const subscribedListeners = listeners.size;
          session.dispose();
          const disposedListeners = listeners.size, disposed = await pending;
          return { bytes: [...saved.bytes], report: saved.report, document, bindings, stale, sourceMatches, subscribedListeners, disposedListeners, disposed };
        } finally { session.dispose(); }
      }, { input: [...input], sourceXml, esm });
      assert.deepEqual(result.bytes, expected.bytes); assert.deepEqual(result.report, expected.report);
      assert.deepEqual(result.document, expected.document); assert.deepEqual(result.bindings, expected.bindings);
      assert.equal(result.stale, "save-snapshot-stale"); assert.equal(result.sourceMatches, true); assert.deepEqual(requests, []);
      assert.equal(result.subscribedListeners, 1); assert.equal(result.disposedListeners, 0); assert.equal(result.disposed, "session-disposed");
      await page.close();
    }
  } finally { await browser.close(); }
});

test("v2 shape/notes saves match Node in browser ESM/global with independent per-part DOM identities", async () => {
  const { createPresentation, normalizePresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const { createTextOwnerFixture } = await import("../tests/fixtures/text-owners.mjs");
  const { readGeneratedZip } = await import("../tests/support/read-generated-zip.mjs");
  const input = (await generatePptx(createTextOwnerFixture(createPresentation))).bytes;
  const profile = { projection: "content", editSaveProfile: "text-runs-v2" }, parts = readGeneratedZip(input);
  const sourceXml = Object.fromEntries([...parts].filter(([name]) => /^(ppt\/slides\/slide|ppt\/notesSlides\/notesSlide)\d+\.xml$/.test(name)).map(([name, bytes]) => ["/" + name, bytes.toString()]));
  const edit = session => {
    const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "shape" && e.text);
    const shape = structuredClone(slide.removeElement(slide.elements[index].id));
    shape.text.content[0].runs[0].text = "Browser shape <&🙂"; slide.insertElement(index, shape);
    const notes = structuredClone(slide.notes); notes[0].runs[0].text = "Browser notes <&🙂"; slide.setNotes(notes);
  };
  const node = await importPptx(input, profile); edit(node);
  const prepared = await node.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
  const saved = await node.save(prepared.plan), expected = { bytes: [...saved.bytes], report: saved.report, document: normalizePresentation(node.document), drawings: node.report.textRuns, notes: node.report.notesTextRuns };
  node.dispose();
  const browser = await chromium.launch();
  try {
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = [];
      await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(bundlePath(name, format), "utf8") });
      const esm = format === "esm" ? {
        core: javascriptDataUrl(await readFile(bundlePath("core", format), "utf8")),
        parser: javascriptDataUrl(await readFile(bundlePath("pptx-parser", format), "utf8")),
      } : undefined;
      const result = await page.evaluate(async ({ input, sourceXml, esm, profile }) => {
        const api = esm ? await import(esm.parser) : PPTKit, core = esm ? await import(esm.core) : PPTKit;
        const session = await api.importPptx(Uint8Array.from(input), profile);
        try {
          const sourceMatches = [...session.report.textRuns, ...session.report.notesTextRuns].every(binding => {
            const dom = new DOMParser().parseFromString(sourceXml[binding.source.part], "application/xml");
            const element = [...dom.getElementsByTagName("*")][binding.sourceElementIndex - 1];
            return element?.localName === "t" && element.namespaceURI.endsWith("/main");
          });
          const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "shape" && e.text);
          const shape = structuredClone(slide.removeElement(slide.elements[index].id));
          shape.text.content[0].runs[0].text = "Browser shape <&🙂"; slide.insertElement(index, shape);
          const notes = structuredClone(slide.notes); notes[0].runs[0].text = "Browser notes <&🙂"; slide.setNotes(notes);
          const prepared = await session.prepareSave({ mode: "edit" });
          if (prepared.status !== "ready") throw new Error(JSON.stringify(prepared.report));
          const saved = await session.save(prepared.plan), document = core.normalizePresentation(session.document);
          const reopened = await api.importPptx(saved.bytes, profile);
          const roundTrip = JSON.stringify(core.normalizePresentation(reopened.document)) === JSON.stringify(document);
          reopened.dispose();
          notes[0].runs[0].text = "stale notes"; slide.setNotes(notes);
          let stale; try { await session.save(prepared.plan); } catch (error) { stale = error.code; }
          return { bytes: [...saved.bytes], report: saved.report, document, drawings: session.report.textRuns, notes: session.report.notesTextRuns, sourceMatches, roundTrip, stale };
        } finally { session.dispose(); }
      }, { input: [...input], sourceXml, esm, profile });
      const { sourceMatches, roundTrip, stale, ...actual } = result;
      assert.deepEqual(actual, expected); assert.equal(sourceMatches, true); assert.equal(roundTrip, true);
      assert.equal(stale, "save-snapshot-stale"); assert.deepEqual(requests, []);
      await page.close();
    }
  } finally { await browser.close(); }
});

test("v3 table saves match Node in browser ESM/global and independently locate physical anchor columns", async () => {
  const { createPresentation, normalizePresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  const { createTableSaveFixture } = await import("../tests/fixtures/table-save.mjs");
  const { readGeneratedZip } = await import("../tests/support/read-generated-zip.mjs");
  const input = (await generatePptx(createTableSaveFixture(createPresentation))).bytes;
  const profile = { projection: "content", editSaveProfile: "text-runs-v3" };
  const sourceXml = readGeneratedZip(input).get("ppt/slides/slide1.xml").toString();
  const node = await importPptx(input, profile), slide = node.document.slides[0];
  const original = slide.elements.find(e => e.type === "table" && e.columns.length === 3), index = slide.elements.indexOf(original), table = structuredClone(original);
  table.rows[1].cells[0].content[0].runs[0].text = "Browser merged <&🙂";
  table.rows[2].cells[0].content[0].runs[0].text = "Browser last row";
  slide.removeElement(original.id); slide.insertElement(index, table);
  const plan = await node.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready");
  const saved = await node.save(plan.plan), expected = { bytes: [...saved.bytes], report: saved.report, document: normalizePresentation(node.document), bindings: node.report.tableTextRuns };
  node.dispose();
  const browser = await chromium.launch();
  try {
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = [];
      await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(bundlePath(name, format), "utf8") });
      const esm = format === "esm" ? {
        core: javascriptDataUrl(await readFile(bundlePath("core", format), "utf8")),
        parser: javascriptDataUrl(await readFile(bundlePath("pptx-parser", format), "utf8")),
      } : undefined;
      const result = await page.evaluate(async ({ input, sourceXml, profile, esm }) => {
        const api = esm ? await import(esm.parser) : PPTKit, core = esm ? await import(esm.core) : PPTKit;
        const session = await api.importPptx(Uint8Array.from(input), profile);
        try {
          const dom = new DOMParser().parseFromString(sourceXml, "application/xml"), elements = [...dom.getElementsByTagName("*")];
          const sourceMatches = session.report.tableTextRuns.every(binding => {
            const atom = elements[binding.sourceElementIndex - 1]; let cell = atom;
            while (cell && cell.localName !== "tc") cell = cell.parentElement;
            if (!cell || atom.localName !== "t") return false;
            const row = cell.parentElement;
            return [...row.children].filter(n => n.localName === "tc").indexOf(cell) === binding.columnIndex &&
              [...row.parentElement.children].filter(n => n.localName === "tr").indexOf(row) === binding.rowIndex;
          });
          const slide = session.document.slides[0], original = slide.elements.find(e => e.type === "table" && e.columns.length === 3), index = slide.elements.indexOf(original), table = structuredClone(original);
          table.rows[1].cells[0].content[0].runs[0].text = "Browser merged <&🙂";
          table.rows[2].cells[0].content[0].runs[0].text = "Browser last row";
          slide.removeElement(original.id); slide.insertElement(index, table);
          const plan = await session.prepareSave({ mode: "edit" });
          if (plan.status !== "ready") throw new Error(JSON.stringify(plan.report));
          if (session.assertSavePlanCurrent(plan.plan) !== undefined) throw new Error("Plan guard must be synchronous and return no output");
          if ("importPptxFile" in api || "writePptxFile" in api) throw new Error("Node file adapter leaked into a browser entry");
          const saved = await session.save(plan.plan), document = core.normalizePresentation(session.document);
          const reopened = await api.importPptx(saved.bytes, profile);
          const roundTrip = JSON.stringify(core.normalizePresentation(reopened.document)) === JSON.stringify(document);
          reopened.dispose();
          table.rows[1].cells[0].content[0].runs[0].text = "stale";
          slide.removeElement(original.id); slide.insertElement(index, table);
          let stale; try { await session.save(plan.plan); } catch (error) { stale = error.code; }
          try { session.assertSavePlanCurrent(plan.plan); throw new Error("Stale synchronous plan guard succeeded"); } catch (error) { if (error.code !== "save-snapshot-stale") throw error; }
          return { bytes: [...saved.bytes], report: saved.report, document, bindings: session.report.tableTextRuns, sourceMatches, roundTrip, stale };
        } finally { session.dispose(); }
      }, { input: [...input], sourceXml, profile, esm });
      const { sourceMatches, roundTrip, stale, ...actual } = result;
      assert.deepEqual(actual, expected); assert.equal(sourceMatches, true); assert.equal(roundTrip, true);
      assert.equal(stale, "save-snapshot-stale"); assert.deepEqual(requests, []); await page.close();
    }
  } finally { await browser.close(); }
});

test("finite timeline plans and random-access samples agree in Node, browser ESM/global and JSON transport", async t => {
  const api = await import("../packages/core/dist/index.js");
  const models = [
    { kind: "sequence", id: "root", delayMs: 5, children: [{ kind: "span", id: "α😀", durationMs: 10, repeat: 3, fill: "hold" }, { kind: "parallel", id: "parallel", children: [{ kind: "span", id: "zero", durationMs: 0 }, { kind: "span", id: "delayed", delayMs: 4, durationMs: 7 }] }] },
    { kind: "parallel", id: "empty", children: [], delayMs: 6 },
    { kind: "span", id: "huge", durationMs: 1, repeat: Number.MAX_SAFE_INTEGER, fill: "hold" },
  ];
  const times = [0, 5, 14, 15, 35, 39, 46, Number.MAX_SAFE_INTEGER - 1, Number.MAX_SAFE_INTEGER], options = { maxTimeMs: Number.MAX_SAFE_INTEGER, maxRepeat: Number.MAX_SAFE_INTEGER };
  const expected = models.map(model => { const plan = api.compileFiniteTimeline(model, options); return { plan, samples: times.map(at => api.sampleFiniteTimeline(plan, at, options)) }; });
  const browser = await chromium.launch(); t.after(() => browser.close());
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    if (format === "global") await page.addScriptTag({ content: await readFile(bundlePath("core", "global"), "utf8") });
    const esm = format === "esm" ? javascriptDataUrl(await readFile(bundlePath("core", "esm"), "utf8")) : null;
    const result = await page.evaluate(async ({ models, times, options, esm }) => {
      const api = esm ? await import(esm) : PPTKit;
      return models.map(model => {
        const plan = api.compileFiniteTimeline(model, options), transported = JSON.parse(JSON.stringify(plan));
        const samples = times.map(at => api.sampleFiniteTimeline(transported, at, options));
        if (!Object.isFrozen(plan) || !Object.isFrozen(plan.boundaries) || !samples.every(rows => Object.isFrozen(rows) && rows.every(Object.isFrozen))) throw Error("Timeline snapshots must be frozen");
        return { plan, samples };
      });
    }, { models, times, options, esm });
    assert.deepEqual(result, expected); assert.deepEqual(requests, []); await page.close();
  }
});

test("finite timeline browser diagnostics reject unknown semantics and getter/corrupt schedule inputs", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(); await page.addScriptTag({ content: await readFile(bundlePath("core", "global"), "utf8") });
  const result = await page.evaluate(() => {
    let called = false; const codes = [], span = { kind: "span", id: "x", durationMs: 10 };
    const actions = [
      () => PPTKit.compileFiniteTimeline({ ...span, repeat: Infinity }),
      () => PPTKit.compileFiniteTimeline({ ...span, repeat: .5 }),
      () => PPTKit.compileFiniteTimeline({ ...span, autoReverse: true }),
      () => PPTKit.compileFiniteTimeline({ ...span, get durationMs() { called = true; return 10; } }),
      () => { const plan = JSON.parse(JSON.stringify(PPTKit.compileFiniteTimeline(span))); plan.boundaries[1].timeMs = 9; PPTKit.sampleFiniteTimeline(plan, 0); },
    ];
    for (const run of actions) try { run(); codes.push("unexpected-success"); } catch (error) { if (!(error instanceof PPTKit.FiniteTimelineError)) throw error; codes.push(error.code); }
    return { called, codes };
  });
  assert.deepEqual(result, { called: false, codes: ["unsupported-timeline-feature", "unsupported-timeline-feature", "unsupported-timeline-feature", "invalid-timeline-input", "invalid-timeline-plan"] });
});

test("event timeline replay agrees across Node, ESM/global, JSON transport, visits and current inventories", async t => {
  const node = await import("../packages/core/dist/index.js"), input = { pageId: "p", elementIds: ["shape", "button"], stages: [
    { id: "enter", target: { pageId: "p", elementId: "shape" }, trigger: { type: "enter" }, timing: { kind: "span", id: "leaf", durationMs: 5 } },
    { id: "点击😀", target: { pageId: "p", elementId: "shape" }, trigger: { type: "object-click", elementId: "button" }, timing: { kind: "span", id: "leaf", delayMs: 2, durationMs: 3, repeat: 2, fill: "hold" } },
    { id: "after", target: { pageId: "p", elementId: "shape" }, trigger: { type: "after-previous" }, timing: { kind: "span", id: "zero", durationMs: 0 } },
  ] };
  const events = [{ id: "early", pageId: "p", visitId: "visit", timeMs: 1, type: "object-click", elementId: "button" }, { id: "old", pageId: "p", visitId: "old", timeMs: 5, type: "object-click", elementId: "button" }, { id: "now", pageId: "p", visitId: "visit", timeMs: 5, type: "object-click", elementId: "button" }];
  const plan = node.compileEventTimeline(input), times = [0, 1, 5, 7, 8, 10, 13, 100, 5], expected = { plan, replays: times.map(timeMs => node.replayEventTimeline(plan, { visitId: "visit", timeMs, elementIds: input.elementIds, events })), reentry: node.replayEventTimeline(plan, { visitId: "new", timeMs: 0, elementIds: input.elementIds, events }) };
  const browser = await chromium.launch(); t.after(() => browser.close());
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(), requests = [];
    await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    if (format === "global") await page.addScriptTag({ content: await readFile(bundlePath("core", format), "utf8") });
    const esm = format === "esm" ? javascriptDataUrl(await readFile(bundlePath("core", format), "utf8")) : null;
    const actual = await page.evaluate(async ({ input, events, times, esm }) => {
      const api = esm ? await import(esm) : PPTKit, plan = api.compileEventTimeline(input), transported = JSON.parse(JSON.stringify(plan));
      const replays = times.map(timeMs => api.replayEventTimeline(transported, { visitId: "visit", timeMs, elementIds: input.elementIds, events }));
      if (!Object.isFrozen(plan) || !replays.every(r => Object.isFrozen(r) && Object.isFrozen(r.events) && r.stages.every(s => Object.isFrozen(s.samples)))) throw Error("Event outputs must be frozen");
      return { plan, replays, reentry: api.replayEventTimeline(transported, { visitId: "new", timeMs: 0, elementIds: input.elementIds, events }) };
    }, { input, events, times, esm });
    assert.deepEqual(actual, expected); assert.deepEqual(requests, []); await page.close();
  }
});

test("browser event timelines reject deleted targets, duplicate events, corrupt future plans and accessors", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(); await page.addScriptTag({ content: await readFile(bundlePath("core", "global"), "utf8") });
  const result = await page.evaluate(() => {
    const input = { pageId: "p", elementIds: ["s"], stages: [{ id: "a", target: { pageId: "p", elementId: "s" }, trigger: { type: "page-click" }, timing: { kind: "span", id: "t", durationMs: 1 } }] }, plan = PPTKit.compileEventTimeline(input);
    const request = { visitId: "v", timeMs: 0, elementIds: ["s"], events: [] }, event = { id: "e", pageId: "p", visitId: "v", timeMs: 0, type: "page-click" };
    let called = false; const codes = [];
    const actions = [
      () => PPTKit.replayEventTimeline(plan, { ...request, elementIds: [] }),
      () => PPTKit.replayEventTimeline(plan, { ...request, events: [event, event] }),
      () => { const p = JSON.parse(JSON.stringify(plan)); p.stages[0].timing.boundaries.pop(); PPTKit.replayEventTimeline(p, request); },
      () => PPTKit.replayEventTimeline(plan, { ...request, events: [{ ...event, get timeMs() { called = true; return 0; } }] }),
      () => PPTKit.compileEventTimeline({ ...input, stages: [{ ...input.stages[0], trigger: { type: "media-end" } }] }),
    ];
    for (const run of actions) try { run(); codes.push("unexpected-success"); } catch (error) { if (!(error instanceof PPTKit.EventTimelineError)) throw error; codes.push(error.code); }
    return { called, codes };
  });
  assert.deepEqual(result, { called: false, codes: ["invalid-event-timeline-input", "invalid-event-timeline-input", "invalid-event-timeline-plan", "invalid-event-timeline-input", "unsupported-event-timeline-feature"] });
});

test("bound slide timelines survive Core/IR/Layout and copy in browser ESM/global with explicit format loss boundaries", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    const sources = Object.fromEntries(await Promise.all(["core", "layout", "pptx-exporter", "svg-renderer"].map(async name => [name, await readFile(bundlePath(name, format), "utf8")])));
    if (format === "global") for (const source of Object.values(sources)) await page.addScriptTag({ content: source });
    const urls = {};
    if (format === "esm") {
      urls.core = javascriptDataUrl(sources.core);
      urls.layout = javascriptDataUrl(sources.layout.replaceAll('"@pptkit/core"', JSON.stringify(urls.core)));
      urls.exporter = javascriptDataUrl(sources["pptx-exporter"].replaceAll('"@pptkit/core"', JSON.stringify(urls.core)).replaceAll('"@pptkit/layout"', JSON.stringify(urls.layout)));
      urls.renderer = javascriptDataUrl(sources["svg-renderer"].replaceAll('"@pptkit/core"', JSON.stringify(urls.core)).replaceAll('"@pptkit/layout"', JSON.stringify(urls.layout)));
    }
    const result = await page.evaluate(async ({ format, urls }) => {
      const core = format === "esm" ? await import(urls.core) : PPTKit, layout = format === "esm" ? await import(urls.layout) : PPTKit, exporter = format === "esm" ? await import(urls.exporter) : PPTKit, renderer = format === "esm" ? await import(urls.renderer) : PPTKit;
      const doc = core.createPresentation(), slide = doc.addSlide({ id: "p" });
      slide.addElement({ type: "text", id: "t", box: { x: 10, y: 10, width: 200, height: 40 }, content: "Original" });
      const control = await renderer.renderPresentationToSvg(doc);
      slide.setEventTimeline({ stages: [{ id: "s", target: { pageId: "p", elementId: "t" }, trigger: { type: "enter" }, timing: { kind: "span", id: "time", durationMs: 1 } }] });
      let removed; try { slide.removeElement("t"); } catch (error) { removed = error.code; }
      const changed = structuredClone(slide.elements[0]); changed.content = "Changed"; slide.replaceElement("t", changed);
      const copy = doc.duplicateSlide("p"), normalized = core.normalizePresentation(doc), resolved = layout.resolveNormalizedLayout(JSON.parse(JSON.stringify(normalized)));
      const svg = await renderer.renderPresentationToSvg(doc);
      let blocked; try { await exporter.generatePptx(doc); } catch (error) { if (!(error instanceof exporter.PptxExportError)) throw error; blocked = error.code; }
      const omitted = await exporter.generatePptx(doc, { eventTimelineMode: "omit" });
      return { removed, copyTarget: copy.eventTimeline.stages[0].target.elementId === copy.elements[0].id, copyPage: copy.eventTimeline.stages[0].target.pageId === copy.id, retained: JSON.stringify(resolved.slides[0].eventTimeline.plan) === JSON.stringify(slide.getEventTimelinePlan()) && JSON.stringify(resolved.slides[0].eventTimeline.source) === JSON.stringify(slide.eventTimeline), warnings: svg.warnings.map(w => w.code), blocked, omittedWarnings: omitted.warnings.map(w => w.code), bytes: [...omitted.bytes.slice(0, 2)], controlUnaffected: control.warnings.length === 0 };
    }, { format, urls });
    assert.deepEqual(result, { removed: "invalid-event-timeline-input", copyTarget: true, copyPage: true, retained: true, warnings: ["event-timeline-static", "event-timeline-static"], blocked: "event-timeline-export-unsupported", omittedWarnings: ["event-timeline-omitted", "event-timeline-omitted"], bytes: [80, 75], controlUnaffected: true });
    assert.deepEqual(requests, []); await page.close();
  }
});

test('SVG slide selection agrees across Node, browser ESM and global entries', async t => {
  const core = await import('../packages/core/dist/index.js'), renderer = await import('../packages/svg-renderer/dist/index.js');
  const run = async (make, render) => {
    const d = make({ id: 'deck' });
    d.addSlide({ id: 'one', elements: [{ id: 'a', type: 'text', box: { x: 1, y: 1, width: 100, height: 20 }, content: 'First' }] });
    d.addSlide({ id: 'two', hidden: true, elements: [{ id: 'b', type: 'text', box: { x: 1, y: 1, width: 100, height: 20 }, content: 'Second 中😀' }] });
    const result = await render(d, { slideIndices: [1] });
    const bytes = new TextEncoder().encode(result.slides[0].svg).length;
    const exact = await render(d, { slideIndices: [1], maxSvgBytes: bytes });
    if (exact.slides[0].svg !== result.slides[0].svg) throw Error('boundary changed output');
    let limitError;
    try { await render(d, { slideIndices: [1], maxSvgBytes: bytes - 1 }); throw Error('missing limit failure'); }
    catch (error) { if (error.code !== 'svg-limit-exceeded') throw error; limitError = { code: error.code, slideIndex: error.slideIndex }; }
    const controller = new AbortController(); controller.abort(); let abortCode;
    try { await render(d, { signal: controller.signal }); throw Error('missing abort failure'); }
    catch (error) { if (error.code !== 'operation-aborted') throw error; abortCode = error.code; }
    const waiting = make({ id: 'waiting' }); waiting.registerAsset({ id: 'img', kind: 'image', source: { type: 'url', value: 'https://example.invalid/a' }, mimeType: 'image/png', width: 1, height: 1 });
    waiting.addSlide({ id: 'image', elements: [{ id: 'picture', type: 'image', assetId: 'img', box: { x: 0, y: 0, width: 10, height: 10 } }] });
    const live = new AbortController(); let started; const ready = new Promise(resolve => { started = resolve; });
    const pending = render(waiting, { signal: live.signal, resolveAsset: (_asset, context) => { if (context.signal !== live.signal) throw Error('wrong signal'); started(); return new Promise(() => {}); } });
    await ready; live.abort(); let pendingAbort;
    try { await pending; throw Error('missing live abort'); } catch (error) { if (error.code !== 'operation-aborted') throw error; pendingAbort = error.code; }
    return { result, limitError, abortCode, pendingAbort };
  };
  const expected = await run(core.createPresentation, renderer.renderPresentationToSvg);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(bundlePath(name, 'global'), 'utf8') });
  const imports = {};
  for (const name of ['core', 'layout', 'svg-renderer']) imports[`@pptkit/${name}`] = javascriptDataUrl(await readFile(bundlePath(name, 'esm')));
  await page.addScriptTag({ type: 'importmap', content: JSON.stringify({ imports }) });
  const result = await page.evaluate(async ({ coreUrl, rendererUrl, runSource }) => {
    const run = (0, eval)(`(${runSource})`), core = await import(coreUrl), renderer = await import(rendererUrl);
    return Promise.all([run(core.createPresentation, renderer.renderPresentationToSvg), run(PPTKit.createPresentation, PPTKit.renderPresentationToSvg)]);
  }, { coreUrl: javascriptDataUrl(await readFile(bundlePath('core', 'esm'))), rendererUrl: javascriptDataUrl(await readFile(bundlePath('svg-renderer', 'esm'))), runSource: run.toString() });
  for (const actual of result) assert.deepEqual(actual, expected); assert.deepEqual(requests, []);
});

test('atomic asset batches match across Node and browser entries', async t => {
  const nodeCore = await import('../packages/core/dist/index.js');
  const run = create => {
    const d = create({ id: 'batch' });
    const item = { kind: 'image', source: { type: 'bytes', value: new Uint8Array([1, 2, 3]) }, width: 2, height: 2 };
    const assets = d.registerAssets([{ asset: item, options: { deduplicate: 'content' } }, { asset: item, options: { deduplicate: 'content' } }]);
    let blocked = false; try { d.registerAssets([{ asset: { ...item, id: 'a' } }, { asset: { ...item, id: assets[0].id, width: 7 } }]); } catch { blocked = true; }
    const next = d.registerAssets([{ asset: { ...item, width: 3 } }]);
    const other = create({ id: 'source' }); other.registerAssets([{ asset: { ...item, id: 'source-a' } }, { asset: { ...item, id: 'source-b' } }]);
    const copied = d.copyAssetsFrom(other, ['source-a', 'source-b'], { deduplicate: 'content', external: 'reject' });
    return { ids: assets.map(a => a.id), same: assets[0] === assets[1], blocked, next: next[0].id, copied, assets: d.assets };
  };
  const expected = run(nodeCore.createPresentation), browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath('core', 'global'), 'utf8') });
  const result = await page.evaluate(async ({ url, source }) => { const run = (0, eval)(`(${source})`), core = await import(url); return [run(core.createPresentation), run(PPTKit.createPresentation)]; }, { url: javascriptDataUrl(await readFile(bundlePath('core', 'esm'))), source: run.toString() });
  for (const actual of result) assert.deepEqual(actual, expected); assert.deepEqual(requests, []);
});

test('cross-document slide copy matches Node and both browser entries', async t => {
  const nodeCore = await import('../packages/core/dist/index.js');
  const run = create => {
    const source = create({ id: 'source' }), target = create({ id: 'target' });
    source.registerAsset({ id: 'asset', kind: 'image', source: { type: 'bytes', value: new Uint8Array([1, 2]) } });
    source.defineSlideMaster({ id: 'master', name: 'Master', theme: { colors: { accent1: '123456' } }, elements: [{ id: 'decoration', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 } }] });
    source.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', elements: [{ id: 'layout-shape', type: 'shape', shape: 'rect', box: { x: 20, y: 0, width: 10, height: 10 } }, { id: 'layout-edge', type: 'connector', start: { elementId: 'layout-shape' }, end: { x: 60, y: 10 } }], placeholders: [{ key: 'body', kind: 'body', box: { x: 0, y: 0, width: 100, height: 40 } }] });
    source.addSlide({ id: 'page', elements: [{ id: 'image', type: 'image', assetId: 'asset', box: { x: 0, y: 0, width: 10, height: 10 }, action: { type: 'slide', slideId: 'page' } }] });
    const assets = { deduplicate: 'content', external: 'reject' };
    const master = target.copySlideMasterFrom(source, 'master', { id: 'copied-master', theme: 'source-effective', assets });
    const reusedMaster = target.copySlideMasterFrom(source, 'master', { id: 'unused-master', theme: 'source-effective', assets, reuseEquivalent: true });
    if (reusedMaster.master.id !== 'copied-master' || reusedMaster.elementIds.decoration !== master.elementIds.decoration) throw new Error('Master was not reused');
    const layout = target.copySlideLayoutFrom(source, 'layout', { id: 'copied-layout', masterId: 'copied-master', theme: 'destination', assets });
    const reused = target.copySlideLayoutFrom(source, 'layout', { id: 'unused-layout', masterId: 'copied-master', theme: 'destination', assets, reuseEquivalent: true });
    if (reused.layout.id !== 'copied-layout' || reused.elementIds['layout-shape'] !== layout.elementIds['layout-shape'] || target.layouts.length !== 1) throw new Error('Layout was not reused');
    const result = target.copySlideFrom(source, 'page', { layoutId: 'copied-layout', theme: 'destination', geometry: 'preserve', customData: 'preserve', assets });
    source.addSlide({ id: 'next', elements: [{ id: 'text', type: 'text', content: 'Next', box: { x: 0, y: 0, width: 100, height: 30 }, action: { type: 'slide', slideId: 'page' } }] });
    const batch = target.copySlidesFrom(source, ['next', 'page'], { layoutId: 'copied-layout', theme: 'destination', geometry: 'preserve', customData: 'preserve', assets });
    source.removeSlide('page');
    return { id: result.slide.id, elements: result.slide.elements, elementIds: result.elementIds, mappings: result.assets, assets: target.assets, master: master.master, layout: layout.layout, reused: reused.layout.id, batch: { ids: batch.slideIds, elements: batch.elementIds, action: batch.slides[0].elements[0].action, assets: batch.assets } };
  };
  const expected = run(nodeCore.createPresentation), browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath('core', 'global'), 'utf8') });
  const results = await page.evaluate(async ({ url, source }) => { const run = (0, eval)(`(${source})`), core = await import(url); return [run(core.createPresentation), run(PPTKit.createPresentation)]; }, { url: javascriptDataUrl(await readFile(bundlePath('core', 'esm'))), source: run.toString() });
  for (const actual of results) assert.deepEqual(actual, expected); assert.deepEqual(requests, []);
});

test('per-page batch copy mappings and geometry fit agree across Node and browser entries', async t => {
  const nodeCore = await import('../packages/core/dist/index.js');
  const run = create => {
    const source = create({ size: { width: 400, height: 200 } }), target = create({ size: { width: 200, height: 200 } }), box = { x: 0, y: 0, width: 100, height: 30 };
    source.defineSlideLayout({ id: 'source-layout', name: 'Source', placeholders: [{ key: 'body', kind: 'body', index: 0, box }] });
    for (const id of ['a', 'b']) {
      source.addSlide({ id, layoutId: 'source-layout', elements: [{ id: `text-${id}`, type: 'text', box, content: id, placeholderKey: 'body' }] });
      target.defineSlideLayout({ id: `layout-${id}`, name: id, placeholders: [{ key: `body-${id}`, kind: 'body', index: 0, box }] });
    }
    const r = target.copySlidesFrom(source, ['a', 'b'], { layoutId: null, layouts: { a: 'layout-a', b: 'layout-b' }, theme: 'destination', geometry: 'fit-explicit-geometry', customData: 'preserve', placeholderRoles: 'require-match-and-index', placeholderMatching: 'explicit-index', assets: { deduplicate: 'content', external: 'reject' } });
    source.addSlide({ id: 'opaque', customData: { reference: 'a' } });
    let rejected = false;
    try { target.copySlideFrom(source, 'opaque', { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'reject', assets: { deduplicate: 'content', external: 'reject' } }); } catch (error) { rejected = /opaque custom data/.test(error.message); }
    if (!rejected || target.slides.length !== 2) throw new Error('Custom-data rejection must be atomic');
    const mapping = { body: 'body-a' };
    const changing = Object.create(source, { assets: { get() { mapping.body = 'body-b'; return source.assets; } } });
    const captured = target.copySlideFrom(changing, 'a', { layoutId: 'layout-a', theme: 'destination', geometry: 'preserve', customData: 'reject', placeholders: mapping, assets: { deduplicate: 'content', external: 'reject' } });
    if (mapping.body !== 'body-b' || captured.slide.elements[0].placeholderKey !== 'body-a') throw new Error('Mapping snapshot changed during source capture');
    return r.slides.map(slide => ({ layoutId: slide.layoutId, placeholderKey: slide.elements[0].placeholderKey, box: slide.elements[0].box }));
  };
  const expected = run(nodeCore.createPresentation), browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  assert.deepEqual(expected[0].box, { x: 0, y: 50, width: 50, height: 15 });
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(bundlePath('core', 'global'), 'utf8') });
  const results = await page.evaluate(async ({ url, source }) => { const run = (0, eval)(`(${source})`), core = await import(url); return [run(core.createPresentation), run(PPTKit.createPresentation)]; }, { url: javascriptDataUrl(await readFile(bundlePath('core', 'esm'))), source: run.toString() });
  for (const actual of results) assert.deepEqual(actual, expected); assert.deepEqual(requests, []);
});

test("strict URL image loading and byte deduplication match Chromium and Node", async t => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/node.js");
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  const still = await readFile(new URL("../tests/fixtures/transparent-replacement.png", import.meta.url));
  const corrupt = Buffer.from(still);
  corrupt[corrupt.length - 1] ^= 1;
  const options = { externalImageValidation: "signature", pngDiagnostics: "reject-scanlines", deduplicateResolvedAssets: true };
  const args = { bytes: [...still], base64: still.toString("base64"), corruptBase64: corrupt.toString("base64"), options };
  const build = async ({ bytes, base64, corruptBase64, options, api = globalThis.PPTKit }) => {
    const deck = api.createPresentation();
    deck.registerAsset({ id: "inline", kind: "image", source: { type: "bytes", value: Uint8Array.from(bytes) } });
    for (const [id, mime] of [["valid", "application/octet-stream"], ["mismatch", "image/gif"]]) {
      deck.registerAsset({ id, kind: "image", source: { type: "url", value: `data:${mime};base64,${base64}` } });
    }
    deck.registerAsset({ id: "corrupt", kind: "image", source: { type: "url", value: `data:image/png;base64,${corruptBase64}` } });
    // 1x1 grayscale, valid CRCs and zlib framing, but forbidden scanline filter 5.
    deck.registerAsset({ id: "filter", kind: "image", source: {
      type: "url", value: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGMVBAAAHQAXjsBI9wAAAABJRU5ErkJggg==",
    } });
    deck.addSlide({ elements: ["inline", "valid", "mismatch", "corrupt", "filter"].map((assetId, i) => ({
      type: "image", assetId, box: { x: 20 * i, y: 0, width: 20, height: 20 },
    })) });
    const result = await api.generatePptx(deck, options);
    return { bytes: [...result.bytes], warnings: result.warnings };
  };
  const expected = await build({ ...args, api: { createPresentation, generatePptx } });
  const actual = await page.evaluate(build, args);
  assert.deepEqual(actual, expected);
  assert.deepEqual(expected.warnings.map(w => [w.code, w.assetId]), [
    ["asset-read-failed", "mismatch"], ["asset-read-failed", "corrupt"], ["asset-read-failed", "filter"],
    ["image-omitted", "mismatch"], ["image-omitted", "corrupt"], ["image-omitted", "filter"],
  ]);
  assert.match(expected.warnings.find(w => w.assetId === "filter").message, /invalid row filter/);
  assert.equal([...readGeneratedZip(Uint8Array.from(expected.bytes)).keys()].filter(name => name.startsWith("ppt/media/")).length, 1);
});

test("JPEG marker validation preserves progressive data and omits a truncated URL in Chromium and Node", async t => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/node.js");
  const bytes = await readFile(new URL("../tests/fixtures/jpeg-progressive-components.jpg", import.meta.url));
  const args = { valid: bytes.toString("base64"), truncated: bytes.subarray(0, -2).toString("base64") };
  const build = async ({ valid, truncated, api = globalThis.PPTKit }) => {
    const deck = api.createPresentation();
    for (const [id, data] of Object.entries({ valid, truncated })) {
      deck.registerAsset({ id, kind: "image", source: { type: "url", value: `data:image/jpeg;base64,${data}` } });
    }
    deck.addSlide({ elements: ["valid", "truncated"].map((assetId, i) => ({
      type: "image", assetId, box: { x: i * 20, y: 0, width: 20, height: 20 },
    })) });
    const result = await api.generatePptx(deck, { externalImageValidation: "signature", jpegValidation: "structure" });
    return { bytes: [...result.bytes], warnings: result.warnings };
  };
  const expected = await build({ ...args, api: { createPresentation, generatePptx } });
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(bundlePath(name, "global"), "utf8") });
  assert.deepEqual(await page.evaluate(build, args), expected);
  assert.deepEqual(expected.warnings.map(w => [w.code, w.assetId]), [
    ["asset-read-failed", "truncated"], ["image-omitted", "truncated"],
  ]);
  assert.match(expected.warnings[0].message, /missing EOI/);
  assert.deepEqual(readGeneratedZip(Uint8Array.from(expected.bytes)).get("ppt/media/valid.jpg"), bytes);
});
