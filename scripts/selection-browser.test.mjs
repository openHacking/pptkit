import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation } from '../packages/core/dist/index.js';
import { resolveNormalizedLayout } from '../packages/layout/dist/index.js';
import { getSvgSelectionFrames, getSvgSlideSelectionFrames, getSvgSlideSelectionGeometry, captureSvgPreviewSnapshot, planSvgPreviewInvalidation, renderLayoutToSvg } from '../packages/svg-renderer/dist/index.js';

test('unpainted connector filtering retains independently visible transparent-stroke markers', async t => {
  const doc = createPresentation({ size: { width: 100, height: 100 } });
  const connector = (id, y, style) => ({ type: 'connector', id, start: { x: 20, y }, end: { x: 80, y }, style });
  doc.addSlide({ id: 'slide', elements: [
    connector('none', 10, { paint: { type: 'none' }, width: 4, endArrow: 'triangle' }),
    connector('clear', 30, { paint: { type: 'solid', color: '000000', opacity: 0 }, width: 4 }),
    connector('clear-arrow', 50, { paint: { type: 'solid', color: '000000', opacity: 0 }, width: 4, endArrow: 'triangle' }),
    connector('ink', 70, { paint: { type: 'solid', color: '000000' }, width: 4 }),
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const actual = await page.evaluate(async ({ layout, svg }) => {
    const items = PPTKit.getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedConnectors: true });
    const image = new Image(); image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`; await image.decode();
    const canvas = document.createElement('canvas'); canvas.width = 100; canvas.height = 100;
    const context = canvas.getContext('2d'); context.drawImage(image, 0, 0, 100, 100);
    const pixel = (x, y) => Array.from(context.getImageData(x, y, 1, 1).data);
    return { items, noLine: pixel(40, 10), clearLine: pixel(40, 30), clearArrowLine: pixel(40, 50), arrow: pixel(75, 50), ink: pixel(40, 70) };
  }, { layout, svg });
  assert.deepEqual(actual.items, getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedConnectors: true }));
  assert.deepEqual(actual.items.map(item => item.elementId), ['clear-arrow', 'ink']);
  for (const pixel of [actual.noLine, actual.clearLine, actual.clearArrowLine]) assert.deepEqual(pixel.slice(0, 3), [255, 255, 255]);
  for (const pixel of [actual.arrow, actual.ink]) assert.ok(pixel[0] < 30 && pixel[1] < 30 && pixel[2] < 30);
});

test('large layout diagnostic batch stays ordered in Chromium', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ['core', 'layout']) await page.addScriptTag({ path: `packages/${name}/dist/browser/global.js` });
  const actual = await page.evaluate(() => {
    const document = PPTKit.createPresentation(); document.addSlide({ id: 'slide' });
    const rule = { type: 'align', container: { kind: 'slide', id: 'missing' }, elementIds: [], horizontal: 'start' };
    const diagnostics = PPTKit.resolveLayout(document, { composition: Array(150000).fill(rule) }).analysis.diagnostics;
    return { count: diagnostics.length, first: diagnostics[0].path, last: diagnostics.at(-1).path, code: diagnostics[0].code };
  });
  assert.deepEqual(actual, { count: 150000, first: 'options.composition.0', last: 'options.composition.149999', code: 'composition-container-missing' });
});

test('large connector geometry and composition resolve in Chromium without argument overflow', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ['core', 'layout']) await page.addScriptTag({ path: `packages/${name}/dist/browser/global.js` });
  const result = await page.evaluate(() => {
    const document = PPTKit.createPresentation({ size: { width: 500, height: 300 } });
    const route = Array.from({ length: 150000 }, (_, index) => ({ x: index % 101, y: index % 51 }));
    document.addSlide({ id: 'slide', elements: [{ id: 'line', type: 'connector', start: { x: 0, y: 0 }, end: { x: 100, y: 50 }, route }] });
    const line = PPTKit.resolveLayout(document, { composition: [{ type: 'align', container: { kind: 'slide', id: 'slide' }, elementIds: ['line'], horizontal: 'end' }] }).slides[0].elements[0];
    return { box: line.box, start: line.start, end: line.end, count: line.route.length };
  });
  assert.deepEqual(result, { box: { x: 400, y: 0, width: 100, height: 50 }, start: { x: 400, y: 0 }, end: { x: 500, y: 50 }, count: 150000 });
});

test('browser frame hits handle extreme finite scales and reject unrepresentable differences', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const actual = await page.evaluate(() => {
    const result = [];
    for (const size of [1e-200, 1e200]) {
      const frame = { elementId: 'target', ancestorIds: [], corners: [{ x: 0, y: 0 }, { x: size, y: 0 }, { x: size, y: size }, { x: 0, y: size }] };
      result.push(PPTKit.hitTestSvgSelectionFrames([frame], { x: size / 2, y: size / 2 })?.elementId ?? null);
      result.push(PPTKit.hitTestSvgSelectionFrames([frame], { x: size * 1.5, y: size / 2 })?.elementId ?? null);
    }
    const frame = { elementId: 'overflow', ancestorIds: [], corners: [{ x: -Number.MAX_VALUE, y: 0 }, { x: 0, y: 0 }, { x: 0, y: 1 }, { x: -Number.MAX_VALUE, y: 1 }] };
    try { PPTKit.hitTestSvgSelectionFrames([frame], { x: Number.MAX_VALUE, y: 0 }); result.push('no error'); }
    catch (error) { result.push(error.name); }
    return result;
  });
  assert.deepEqual(actual, ['target', null, 'target', null, 'RangeError']);
});

test('selection frames agree with Chromium SVG transforms and global bundle', async t => {
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'group', id: 'outer', box: { x: 100, y: 100, width: 200, height: 100 }, coordinateSize: { width: 100, height: 100 }, transform: { rotation: 23, flipV: true }, children: [
    { type: 'group', id: 'inner', box: { x: 10, y: 20, width: 40, height: 20 }, coordinateSize: { width: 80, height: 40 }, transform: { rotation: -17, flipH: true }, children: [
      { type: 'shape', id: 'target', shape: 'rect', box: { x: 5, y: 4, width: 30, height: 10 }, transform: { rotation: 14 } },
    ] },
  ] }] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc)), frames = getSvgSelectionFrames(layout.slides[0]);
  const svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(); await page.setContent(svg);
  await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const actual = await page.evaluate(slide => {
    const rect = document.querySelector('[data-pptkit-element-id="target"] rect');
    const ctm = rect.getCTM(), x = Number(rect.getAttribute('x')), y = Number(rect.getAttribute('y'));
    const width = Number(rect.getAttribute('width')), height = Number(rect.getAttribute('height'));
    const corners = [[x, y], [x + width, y], [x + width, y + height], [x, y + height]].map(([x, y]) => ({ x: ctm.a*x + ctm.c*y + ctm.e, y: ctm.b*x + ctm.d*y + ctm.f }));
    return { corners, bundle: PPTKit.getSvgSelectionFrames(slide) };
  }, layout.slides[0]);
  assert.deepEqual(actual.bundle, frames);
  for (let i = 0; i < 4; i++) {
    assert.ok(Math.abs(actual.corners[i].x - frames[0].corners[i].x) < 1e-8);
    assert.ok(Math.abs(actual.corners[i].y - frames[0].corners[i].y) < 1e-8);
  }
});

test('viewport mapping agrees with Chromium screen CTM for letterbox and pillarbox', async t => {
  const doc = createPresentation(); doc.addSlide({ elements: [] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(); await page.setContent(`<div style="margin:17px 31px">${svg}</div>`);
  await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  for (const [width, height] of [[300, 300], [600, 100], [480, 270]]) {
    const result = await page.evaluate(({ width, height, slideSize }) => {
      const svg = document.querySelector('svg'); svg.style.width = `${width}px`; svg.style.height = `${height}px`;
      const rect = svg.getBoundingClientRect(), ctm = svg.getScreenCTM();
      const map = PPTKit.createSvgViewportMapping(slideSize, { x: rect.x, y: rect.y, width: rect.width, height: rect.height });
      const slide = { x: slideSize.width * .37, y: slideSize.height * .61 };
      return { point: map.slideToViewport(slide), back: map.viewportToSlide(map.slideToViewport(slide)), slide,
        expected: { x: ctm.a * slide.x + ctm.c * slide.y + ctm.e, y: ctm.b * slide.x + ctm.d * slide.y + ctm.f } };
    }, { width, height, slideSize: layout.size });
    for (const key of ['x', 'y']) {
      assert.ok(Math.abs(result.point[key] - result.expected[key]) < 1e-8);
      assert.ok(Math.abs(result.back[key] - result.slide[key]) < 1e-8);
    }
  }
});

test('affine client mapping matches Chromium for rotated skewed and reflected SVG', async t => {
  const doc = createPresentation(); doc.addSlide({ elements: [] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.setContent(`<div style="margin:30px;transform:rotate(17deg) skewX(21deg) scale(-1.2,.75);transform-origin:top left">${svg}</div>`);
  await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const result = await page.evaluate(() => {
    const matrix = document.querySelector('svg').getScreenCTM();
    const mapping = PPTKit.createSvgAffineMapping(matrix), slide = { x: 231.25, y: 114.75 };
    const expected = new DOMPoint(slide.x, slide.y).matrixTransform(matrix);
    const client = mapping.slideToViewport(slide), roundTrip = mapping.viewportToSlide(client);
    return { expected: { x: expected.x, y: expected.y }, client, slide, roundTrip, matrix: mapping.matrix };
  });
  for (const axis of ['x', 'y']) {
    assert.ok(Math.abs(result.client[axis] - result.expected[axis]) < 1e-8);
    assert.ok(Math.abs(result.roundTrip[axis] - result.slide[axis]) < 1e-8);
  }
  assert.equal(result.matrix.is2D, true);
});

test('owned selection layers follow rendered SVG order and visibility', async t => {
  const shape = id => ({ type: 'shape', id, shape: 'rect', box: { x: 10, y: 10, width: 50, height: 50 } });
  const doc = createPresentation({ masters: [{ id: 'm', name: 'Master', elements: [shape('m-shape')] }] });
  doc.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm', elements: [shape('l-shape')] });
  doc.addSlide({ id: 'a', layoutId: 'l', elements: [shape('a-shape')] });
  doc.addSlide({ id: 'b', layoutId: 'l', showInheritedShapes: false, elements: [shape('b-shape')] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc)), rendered = await renderLayoutToSvg(layout);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  for (let i = 0; i < layout.slides.length; i++) {
    await page.evaluate(svg => { document.body.innerHTML = svg; }, rendered.slides[i].svg);
    const actual = await page.evaluate(({ layout, i }) => ({
      ids: Array.from(document.querySelectorAll('[data-pptkit-element-id]'), node => node.getAttribute('data-pptkit-element-id')),
      frames: PPTKit.getSvgSlideSelectionFrames(layout, i),
    }), { layout, i });
    const frames = getSvgSlideSelectionFrames(layout, i);
    assert.deepEqual(actual.frames, frames);
    assert.deepEqual(actual.ids, frames.map(f => f.elementId));
  }
});

test('zero-opacity exclusion follows rendered element and group opacity in browser', async t => {
  const doc = createPresentation();
  doc.addSlide({ id: 's', elements: [
    { type: 'shape', id: 'visible', shape: 'rect', box: { x: 0, y: 0, width: 50, height: 50 } },
    { type: 'group', id: 'clear-group', opacity: 0, box: { x: 0, y: 0, width: 50, height: 50 }, coordinateSize: { width: 50, height: 50 }, children: [{ type: 'shape', id: 'nested', shape: 'rect', box: { x: 0, y: 0, width: 50, height: 50 } }] },
    { type: 'shape', id: 'clear', opacity: 0, shape: 'ellipse', box: { x: 0, y: 0, width: 50, height: 50 } },
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.setContent(svg); await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const actual = await page.evaluate(layout => {
    const opacity = id => document.querySelector(`[data-pptkit-element-id="${id}"]`)?.getAttribute('opacity');
    const items = PPTKit.getSvgSlideSelectionGeometry(layout, 0, { excludeZeroOpacity: true });
    return { opacity: ['visible', 'clear-group', 'clear'].map(opacity), items, hit: PPTKit.hitTestSvgSelectionOutlines(items, { x: 25, y: 25 }, 0)?.elementId };
  }, layout);
  assert.deepEqual(actual.opacity, ['1', '0', '0']);
  assert.deepEqual(actual.items, getSvgSlideSelectionGeometry(layout, 0, { excludeZeroOpacity: true }));
  assert.equal(actual.hit, 'visible');
});

test('unpainted-shape filtering follows SVG paint attributes and retains outlined shapes', async t => {
  const doc = createPresentation();
  doc.addSlide({ elements: [
    { type: 'shape', id: 'bottom', shape: 'rect', box: { x: 0, y: 0, width: 50, height: 50 }, style: { fill: { type: 'solid', color: 'FFFFFF' } } },
    { type: 'shape', id: 'empty', shape: 'rect', box: { x: 0, y: 0, width: 50, height: 50 }, style: { fill: { type: 'solid', color: 'FF0000', opacity: 0 }, stroke: { paint: { type: 'none' } } } },
    { type: 'shape', id: 'outline', shape: 'rect', box: { x: 70, y: 0, width: 50, height: 50 }, style: { fill: { type: 'none' }, stroke: { paint: { type: 'solid', color: 'FF0000' }, width: 2 } } },
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.setContent(svg); await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const actual = await page.evaluate(layout => {
    const empty = document.querySelector('[data-pptkit-element-id="empty"] rect');
    const outline = document.querySelector('[data-pptkit-element-id="outline"] rect');
    const items = PPTKit.getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedShapes: true });
    return { empty: [empty.getAttribute('fill-opacity'), empty.getAttribute('stroke')], outline: [outline.getAttribute('fill'), outline.getAttribute('stroke-width')], items, hit: PPTKit.hitTestSvgSelectionOutlines(items, { x: 25, y: 25 }, 0)?.elementId };
  }, layout);
  assert.deepEqual(actual.empty, ['0', 'none']); assert.deepEqual(actual.outline, ['none', '2']);
  assert.deepEqual(actual.items, getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedShapes: true }));
  assert.equal(actual.hit, 'bottom');
});

test('connector selection points match transformed SVG polyline', async t => {
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'group', id: 'outer', box: { x: 100, y: 100, width: 200, height: 100 }, coordinateSize: { width: 100, height: 100 }, transform: { rotation: 23, flipV: true }, children: [
    { type: 'connector', id: 'line', start: { x: 10, y: 10 }, route: [{ x: 70, y: 10 }], end: { x: 70, y: 80 }, transform: { rotation: 13, flipH: true } },
  ] }] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc)), expected = getSvgSlideSelectionGeometry(layout, 0);
  const svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.setContent(svg); await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const actual = await page.evaluate(layout => {
    const line = document.querySelector('[data-pptkit-element-id="line"] polyline'), matrix = line.getCTM();
    const points = Array.from(line.points, p => ({ x: matrix.a * p.x + matrix.c * p.y + matrix.e, y: matrix.b * p.x + matrix.d * p.y + matrix.f }));
    return { points, bundle: PPTKit.getSvgSlideSelectionGeometry(layout, 0) };
  }, layout);
  assert.deepEqual(actual.bundle, expected);
  assert.equal(actual.points.length, expected[0].points.length);
  for (let i = 0; i < actual.points.length; i++) for (const axis of ['x', 'y']) assert.ok(Math.abs(actual.points[i][axis] - expected[0].points[i][axis]) < 1e-8);
});

test('preview invalidation agrees between Node and the browser bundle', async t => {
  const doc = createPresentation();
  doc.addSlide({ id: 'a', elements: [{ type: 'shape', id: 'shape-a', shape: 'rect', box: { x: 1, y: 2, width: 30, height: 20 } }] });
  doc.addSlide({ id: 'b', elements: [{ type: 'shape', id: 'shape-b', shape: 'rect', box: { x: 2, y: 3, width: 30, height: 20 } }] });
  const before = resolveNormalizedLayout(normalizePresentation(doc)), after = structuredClone(before);
  after.slides[0].elements[0].box.x += 10;
  const options = { environmentRevision: 'stable' };
  const expected = planSvgPreviewInvalidation(captureSvgPreviewSnapshot(before, options), captureSvgPreviewSnapshot(after, options));
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const actual = await page.evaluate(({ before, after, options }) => PPTKit.planSvgPreviewInvalidation(PPTKit.captureSvgPreviewSnapshot(before, options), PPTKit.captureSvgPreviewSnapshot(after, options)), { before, after, options });
  assert.deepEqual(actual, expected); assert.deepEqual(actual.render.map(item => item.slideId), ['a']);
  const a = await renderLayoutToSvg(before), b = await renderLayoutToSvg(after);
  assert.notEqual(a.slides[0].svg, b.slides[0].svg); assert.equal(a.slides[1].svg, b.slides[1].svg);
});

test('browser incremental renderer reuses untouched page and matches a full render', async t => {
  const doc = createPresentation();
  doc.addSlide({ id: 'a', elements: [{ type: 'shape', id: 'shape-a', shape: 'rect', box: { x: 1, y: 2, width: 30, height: 20 } }] });
  doc.addSlide({ id: 'b', elements: [{ type: 'shape', id: 'shape-b', shape: 'rect', box: { x: 2, y: 3, width: 30, height: 20 } }] });
  const before = resolveNormalizedLayout(normalizePresentation(doc)), after = structuredClone(before);
  after.slides[0].elements[0].box.x += 10;
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const actual = await page.evaluate(async ({ before, after }) => {
    const options = { environmentRevision: 'stable' };
    const first = await PPTKit.renderIncrementalSvg(before, options);
    const next = await PPTKit.renderIncrementalSvg(after, options, first.cache);
    return { rendered: next.renderedSlideIds, result: next.result };
  }, { before, after });
  assert.deepEqual(actual.rendered, ['a']); assert.deepEqual(actual.result, await renderLayoutToSvg(after));
});

test('preset shape outline hits match Chromium SVG fill geometry under nested transforms', async t => {
  const kinds = ['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'arrow', 'chevron', 'parallelogram'];
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'group', id: 'g', box: { x: 100, y: 100, width: 300, height: 200 }, coordinateSize: { width: 200, height: 200 }, transform: { rotation: 23, flipH: true }, children: kinds.map((shape, i) => ({
    type: 'shape', id: `shape-${i}`, shape, box: { x: 10 + i * 15, y: 20, width: 90, height: 60 }, transform: { rotation: 11 }, ...(shape === 'parallelogram' ? { slant: .4 } : {}),
  })) }] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc)), svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.setContent(svg); await page.addScriptTag({ path: 'packages/svg-renderer/dist/browser/global.js' });
  const differences = await page.evaluate(({ layout, count }) => {
    const items = PPTKit.getSvgSlideSelectionGeometry(layout, 0), mismatches = [];
    for (let i = 0; i < count; i++) {
      const item = items.find(item => item.elementId === `shape-${i}`);
      const node = document.querySelector(`[data-pptkit-element-id="shape-${i}"]`), geometry = node.querySelector('rect, ellipse, polygon');
      const box = geometry.getBBox(), matrix = geometry.getCTM();
      for (let x = 0; x < 10; x++) for (let y = 0; y < 10; y++) {
        const px = box.x + box.width * (x + .37) / 10, py = box.y + box.height * (y + .61) / 10;
        const slide = { x: matrix.a * px + matrix.c * py + matrix.e, y: matrix.b * px + matrix.d * py + matrix.f };
        const native = geometry.isPointInFill(new DOMPoint(px, py));
        const predicted = PPTKit.hitTestSvgSelectionOutlines([item], slide, 0) !== undefined;
        if (native !== predicted) mismatches.push({ i, x, y, native, predicted });
      }
    }
    return mismatches;
  }, { layout, count: kinds.length });
  assert.deepEqual(differences, []);
});
