import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveNormalizedLayout } from '@pptkit/layout';
import { getSvgSelectionFrames, getSvgSlideSelectionFrames, getSvgSlideSelectionGeometry, hitTestSvgSelectionFrames, hitTestSvgSelectionGeometry, hitTestSvgSelectionOutlines, hitTestSvgConnectorStrokes, hitTestSvgConnectorMarkers, hitTestSvgShapeStrokes, hitTestSvgSelectionOutlinesAndStrokes, hitTestSvgSelectionPaintAware } from '../dist/index.js';

const box = (x, y, width, height) => ({ x, y, width, height });
const shape = (id, b, extra = {}) => ({ type: 'shape', id, shape: 'rect', box: b, ...extra });
const close = (point, expected) => { assert.ok(Math.abs(point.x - expected.x) < 1e-9); assert.ok(Math.abs(point.y - expected.y) < 1e-9); };

test('nested group scaling, flip and rotation map frames to slide space in SVG paint order', () => {
  const doc = createPresentation();
  doc.addSlide({ id: 'page', elements: [
    shape('under', box(0, 0, 400, 400)),
    { type: 'group', id: 'outer', box: box(100, 100, 200, 100), coordinateSize: { width: 100, height: 100 }, transform: { rotation: 90 }, children: [
      { type: 'group', id: 'inner', box: box(10, 20, 40, 20), coordinateSize: { width: 40, height: 20 }, transform: { flipH: true }, children: [
        shape('child', box(0, 0, 20, 10)), shape('hidden', box(0, 0, 20, 10), { hidden: true }),
      ] },
    ] },
  ] });
  const slide = resolveNormalizedLayout(normalizePresentation(doc)).slides[0];
  const frames = getSvgSelectionFrames(slide);
  assert.deepEqual(frames.map(f => f.elementId), ['under', 'child']);
  assert.deepEqual(frames[1].ancestorIds, ['outer', 'inner']);
  [[230, 150], [230, 110], [220, 110], [220, 150]].forEach((expected, i) => close(frames[1].corners[i], { x: expected[0], y: expected[1] }));
  assert.equal(hitTestSvgSelectionFrames(frames, { x: 225, y: 130 }).elementId, 'child');
  const [a, b] = frames[1].corners;
  assert.equal(hitTestSvgSelectionFrames(frames, { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 }).elementId, 'child');
  assert.equal(hitTestSvgSelectionFrames(frames, { x: 10, y: 10 }).elementId, 'under');
  assert.equal(hitTestSvgSelectionFrames(frames, { x: 500, y: 500 }), undefined);
  assert.ok(Object.isFrozen(frames) && Object.isFrozen(frames[1]) && Object.isFrozen(frames[1].corners[0]));
  slide.elements[0].box.x = 99;
  assert.equal(frames[0].corners[0].x, 0);
});

test('selection excludes hidden ancestors and connectors, and returns topmost frame', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [shape('bottom', box(0, 0, 40, 40)), shape('top', box(10, 10, 20, 20)),
    { type: 'connector', id: 'line', start: { x: 0, y: 0 }, end: { x: 30, y: 30 } },
    { type: 'group', id: 'hidden-group', hidden: true, box: box(0, 0, 50, 50), coordinateSize: { width: 50, height: 50 }, children: [shape('hidden-child', box(0, 0, 50, 50))] },
  ] });
  const frames = getSvgSelectionFrames(resolveNormalizedLayout(normalizePresentation(doc)).slides[0]);
  assert.deepEqual(frames.map(f => f.elementId), ['bottom', 'top']);
  assert.equal(hitTestSvgSelectionFrames(frames, { x: 20, y: 20 }).elementId, 'top');
  assert.throws(() => hitTestSvgSelectionFrames(frames, { x: NaN, y: 0 }), TypeError);
});

test('frame hits remain stable when finite cross products would overflow or underflow', () => {
  for (const size of [1e-200, 1e200]) {
    const frame = { elementId: 'large-or-small', ancestorIds: [], corners: [
      { x: 0, y: 0 }, { x: size, y: 0 }, { x: size, y: size }, { x: 0, y: size },
    ] };
    const geometry = { ...frame, kind: 'frame', owner: { type: 'slide', id: 's' } };
    for (const hit of [point => hitTestSvgSelectionFrames([frame], point), point => hitTestSvgSelectionGeometry([geometry], point, 0), point => hitTestSvgSelectionOutlines([geometry], point, 0)]) {
      assert.equal(hit({ x: size / 2, y: size / 2 }).elementId, 'large-or-small');
      assert.equal(hit({ x: size, y: size / 2 }).elementId, 'large-or-small');
      assert.equal(hit({ x: size * 1.5, y: size / 2 }), undefined);
    }
  }
  const frame = { elementId: 'overflow', ancestorIds: [], corners: [
    { x: -Number.MAX_VALUE, y: 0 }, { x: 0, y: 0 }, { x: 0, y: 1 }, { x: -Number.MAX_VALUE, y: 1 },
  ] };
  assert.throws(() => hitTestSvgSelectionFrames([frame], { x: Number.MAX_VALUE, y: 0 }), RangeError);
});

test('malformed or unbounded detached group geometry is rejected', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'group', id: 'group', box: box(0, 0, 10, 10), coordinateSize: { width: 10, height: 10 }, children: [shape('child', box(0, 0, 10, 10))] }] });
  const slide = resolveNormalizedLayout(normalizePresentation(doc)).slides[0];
  slide.elements[0].coordinateSize.width = 0;
  assert.throws(() => getSvgSelectionFrames(slide), TypeError);
  slide.elements[0].coordinateSize.width = 10;
  slide.elements[0].box.width = 0;
  assert.deepEqual(getSvgSelectionFrames(slide), []);
  slide.elements[0].box.width = 10;
  slide.elements[0].children.push(slide.elements[0]);
  assert.throws(() => getSvgSelectionFrames(slide), RangeError);
});

test('resolved slide selection observes master, layout, local layer order and visibility', () => {
  const doc = createPresentation({ masters: [{ id: 'master', name: 'Master', elements: [shape('master-item', box(0, 0, 100, 100))] }] });
  doc.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', elements: [shape('layout-item', box(10, 10, 80, 80))] });
  doc.defineSlideLayout({ id: 'hide-master', name: 'Hide master', masterId: 'master', showMasterShapes: false, elements: [shape('layout-2', box(10, 10, 80, 80))] });
  doc.addSlide({ id: 'one', layoutId: 'layout', elements: [shape('local-item', box(20, 20, 60, 60))] });
  doc.addSlide({ id: 'two', layoutId: 'layout', showInheritedShapes: false });
  doc.addSlide({ id: 'three', layoutId: 'hide-master', showInheritedShapes: true });
  doc.addSlide({ id: 'four', layoutId: 'layout' });
  const layout = resolveNormalizedLayout(normalizePresentation(doc)), frames = getSvgSlideSelectionFrames(layout, 0);
  assert.deepEqual(frames.map(f => [f.elementId, f.owner.type, f.owner.id]), [['master-item', 'master', 'master'], ['layout-item', 'layout', 'layout'], ['local-item', 'slide', 'one']]);
  assert.equal(hitTestSvgSelectionFrames(frames, { x: 50, y: 50 }).elementId, 'local-item');
  assert.equal(hitTestSvgSelectionFrames(frames, { x: 15, y: 15 }).elementId, 'layout-item');
  assert.equal(hitTestSvgSelectionFrames(frames, { x: 5, y: 5 }).elementId, 'master-item');
  assert.deepEqual(getSvgSlideSelectionFrames(layout, 1), []);
  assert.deepEqual(getSvgSlideSelectionFrames(layout, 2).map(f => f.elementId), ['layout-2']);
  assert.deepEqual(getSvgSlideSelectionFrames(layout, 3).map(f => f.owner.id), ['master', 'layout']);
  assert.ok(Object.isFrozen(frames[0].owner));
  for (const index of [-1, 4, 1.5, NaN]) assert.throws(() => getSvgSlideSelectionFrames(layout, index), RangeError);
  const broken = structuredClone(layout); delete broken.layouts.find(item => item.id === 'layout').theme;
  assert.throws(() => getSvgSlideSelectionFrames(broken, 0), TypeError);
});

test('connector routes use nested transforms and compete with frames in paint order', () => {
  const doc = createPresentation();
  doc.addSlide({ id: 'page', elements: [shape('bottom', box(0, 0, 400, 400)),
    { type: 'group', id: 'group', box: box(100, 100, 200, 100), coordinateSize: { width: 100, height: 100 }, transform: { flipH: true }, children: [
      { type: 'connector', id: 'line', start: { x: 0, y: 0 }, route: [{ x: 50, y: 0 }], end: { x: 50, y: 50 } },
    ] },
    shape('top', box(190, 115, 20, 20)),
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc)), items = getSvgSlideSelectionGeometry(layout, 0);
  assert.deepEqual(items.map(x => [x.kind, x.elementId]), [['frame', 'bottom'], ['connector', 'line'], ['frame', 'top']]);
  assert.deepEqual(items[1].points, [{ x: 300, y: 100 }, { x: 200, y: 100 }, { x: 200, y: 150 }]);
  assert.deepEqual(items[1].ancestorIds, ['group']); assert.deepEqual(items[1].owner, { type: 'slide', id: 'page' });
  assert.equal(hitTestSvgSelectionGeometry(items, { x: 250, y: 103 }, 3).elementId, 'line');
  assert.equal(hitTestSvgSelectionGeometry(items, { x: 250, y: 104 }, 3).elementId, 'bottom');
  assert.equal(hitTestSvgSelectionGeometry(items, { x: 200, y: 120 }, 1).elementId, 'top');
  assert.equal(hitTestSvgSelectionGeometry(items, { x: 200, y: 145 }, 0).elementId, 'line');
  assert.equal(hitTestSvgSelectionGeometry(items, { x: 600, y: 600 }, 3), undefined);
  assert.ok(Object.isFrozen(items[1].points) && Object.isFrozen(items[1].points[0]));
  for (const tolerance of [-1, NaN, Infinity]) assert.throws(() => hitTestSvgSelectionGeometry(items, { x: 0, y: 0 }, tolerance), RangeError);
  assert.throws(() => hitTestSvgSelectionGeometry(items, { x: NaN, y: 0 }, 1), TypeError);
});

test('optional zero-opacity exclusion applies across owners, leaves, connectors and ancestor groups', () => {
  const doc = createPresentation({ masters: [{ id: 'master', name: 'Master', elements: [shape('master-zero', box(0, 0, 100, 100), { opacity: 0 })] }] });
  doc.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', elements: [shape('layout-visible', box(0, 0, 100, 100))] });
  doc.addSlide({ id: 'page', layoutId: 'layout', elements: [
    shape('bottom', box(0, 0, 100, 100)),
    { type: 'group', id: 'clear-group', opacity: 0, box: box(0, 0, 100, 100), coordinateSize: { width: 100, height: 100 }, children: [shape('group-child', box(0, 0, 100, 100))] },
    { type: 'connector', id: 'clear-line', opacity: 0, start: { x: 0, y: 50 }, end: { x: 100, y: 50 } },
    shape('clear-top', box(0, 0, 100, 100), { opacity: 0 }),
    shape('faint', box(200, 0, 50, 50), { opacity: 0.000001 }),
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const defaults = getSvgSlideSelectionGeometry(layout, 0);
  const filtered = getSvgSlideSelectionGeometry(layout, 0, { excludeZeroOpacity: true });
  assert.deepEqual(defaults.map(item => item.elementId), ['master-zero', 'layout-visible', 'bottom', 'group-child', 'clear-line', 'clear-top', 'faint']);
  assert.deepEqual(filtered.map(item => item.elementId), ['layout-visible', 'bottom', 'faint']);
  assert.equal(hitTestSvgSelectionGeometry(defaults, { x: 50, y: 50 }, 0).elementId, 'clear-top');
  assert.equal(hitTestSvgSelectionOutlines(filtered, { x: 50, y: 50 }, 0).elementId, 'bottom');
  assert.deepEqual(getSvgSlideSelectionGeometry(layout, 0, { excludeZeroOpacity: false }), defaults);
  for (const options of [null, 3, { excludeZeroOpacity: 'true' }]) assert.throws(() => getSvgSlideSelectionGeometry(layout, 0, options), TypeError);
  layout.slides[0].elements[1].opacity = 1;
  assert.deepEqual(filtered.map(item => item.elementId), ['layout-visible', 'bottom', 'faint']);
  assert.ok(getSvgSlideSelectionGeometry(layout, 0, { excludeZeroOpacity: true }).some(item => item.elementId === 'group-child'));
});

test('optional unpainted-shape exclusion is conservative for gradients, patterns, strokes, text and effects', () => {
  const none = { type: 'none' }, clear = { type: 'solid', color: 'FF0000', opacity: 0 }, ink = { type: 'solid', color: 'FF0000' };
  const doc = createPresentation();
  doc.addSlide({ id: 'page', elements: [
    shape('bottom', box(0, 0, 100, 100), { style: { fill: ink } }),
    shape('none', box(0, 0, 100, 100), { style: { fill: none, stroke: { paint: none } } }),
    shape('clear-solid', box(0, 0, 100, 100), { style: { fill: clear, stroke: { paint: ink, width: 0 } } }),
    shape('clear-gradient', box(0, 0, 100, 100), { style: { fill: { type: 'linearGradient', angle: 0, stops: [{ position: 0, color: 'FF0000', opacity: 0 }, { position: 1, color: '0000FF', opacity: 0 }] }, stroke: { paint: none } } }),
    shape('clear-pattern', box(0, 0, 100, 100), { style: { fill: { type: 'pattern', pattern: 'cross', foreground: 'FF0000', background: '0000FF', foregroundOpacity: 0, backgroundOpacity: 0 }, stroke: { paint: none } } }),
    shape('outline', box(200, 0, 100, 100), { style: { fill: none, stroke: { paint: ink, width: 1 } } }),
    shape('text', box(300, 0, 100, 100), { style: { fill: none, stroke: { paint: none } }, text: { content: 'Visible' } }),
    shape('shadow', box(400, 0, 100, 100), { style: { fill: none, stroke: { paint: none }, shadow: { color: '000000', blur: 1, distance: 2, direction: 0 } } }),
    shape('gradient', box(500, 0, 100, 100), { style: { fill: { type: 'linearGradient', angle: 0, stops: [{ position: 0, color: 'FF0000', opacity: 0 }, { position: 1, color: '0000FF', opacity: 1 }] }, stroke: { paint: none } } }),
    shape('pattern', box(600, 0, 100, 100), { style: { fill: { type: 'pattern', pattern: 'cross', foreground: 'FF0000', background: '0000FF', foregroundOpacity: 1, backgroundOpacity: 0 }, stroke: { paint: none } } }),
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const all = getSvgSlideSelectionGeometry(layout, 0), filtered = getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedShapes: true });
  assert.equal(all.length, 10);
  assert.deepEqual(filtered.map(item => item.elementId), ['bottom', 'outline', 'text', 'shadow', 'gradient', 'pattern']);
  assert.equal(hitTestSvgSelectionOutlines(all, { x: 50, y: 50 }, 0).elementId, 'clear-pattern');
  assert.equal(hitTestSvgSelectionOutlines(filtered, { x: 50, y: 50 }, 0).elementId, 'bottom');
  assert.throws(() => getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedShapes: 1 }), TypeError);
});

test('optional unpainted-connector exclusion preserves possible SVG markers', () => {
  const line = (id, style) => ({ type: 'connector', id, start: { x: 0, y: 50 }, end: { x: 100, y: 50 }, style });
  const none = { type: 'none' }, clear = { type: 'solid', color: 'FF0000', opacity: 0 }, ink = { type: 'solid', color: 'FF0000' };
  const doc = createPresentation({ masters: [{ id: 'master', name: 'Master', elements: [line('master-none', { paint: none })] }] });
  doc.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', elements: [line('layout-ink', { paint: ink })] });
  doc.addSlide({ id: 'page', layoutId: 'layout', elements: [
    shape('bottom', box(0, 0, 100, 100)),
    { type: 'group', id: 'group', box: box(0, 0, 100, 100), coordinateSize: { width: 100, height: 100 }, children: [line('group-clear', { paint: clear })] },
    line('none-arrow', { paint: none, endArrow: 'triangle' }),
    line('zero-width', { paint: ink, width: 0 }),
    line('clear', { paint: clear }),
    line('clear-arrow', { paint: clear, endArrow: 'triangle' }),
    line('zero-arrow', { paint: ink, width: 0, beginArrow: 'triangle' }),
    line('faint', { paint: { ...clear, opacity: 0.000001 } }),
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const all = getSvgSlideSelectionGeometry(layout, 0);
  const filtered = getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedConnectors: true });
  assert.deepEqual(filtered.map(item => item.elementId), ['layout-ink', 'bottom', 'clear-arrow', 'zero-arrow', 'faint']);
  assert.deepEqual(getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedConnectors: true, excludeUnpaintedShapes: true, excludeZeroOpacity: true }).map(item => item.elementId), ['layout-ink', 'clear-arrow', 'zero-arrow', 'faint']);
  assert.deepEqual(getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedConnectors: false }), all);
  assert.equal(hitTestSvgSelectionGeometry(filtered, { x: 50, y: 50 }, 0).elementId, 'faint');
  assert.throws(() => getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedConnectors: 'true' }), TypeError);
  layout.slides[0].elements.at(-1).style.paint.opacity = 0;
  assert.equal(filtered.at(-1).elementId, 'faint');
  assert.ok(!getSvgSlideSelectionGeometry(layout, 0, { excludeUnpaintedConnectors: true }).some(item => item.elementId === 'faint'));
});

test('solid connector stroke hit tests use local widths through nonuniform group transforms', () => {
  const doc = createPresentation();
  const line = (id, y, extra = {}) => ({ type: 'connector', id, start: { x: 10, y }, end: { x: 90, y }, style: { paint: { type: 'solid', color: '000000' }, width: 10, ...extra } });
  doc.addSlide({ elements: [{ type: 'group', id: 'group', box: box(100, 100, 200, 100), coordinateSize: { width: 100, height: 100 }, children: [line('line', 50), line('dashed', 70, { dash: 'dash' }), line('clear', 80, { paint: { type: 'solid', color: '000000', opacity: 0 } })] }] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const ordinary = getSvgSlideSelectionGeometry(layout, 0);
  assert.equal(hitTestSvgConnectorStrokes(ordinary, { x: 200, y: 150 }), undefined);
  const items = getSvgSlideSelectionGeometry(layout, 0, { includeConnectorStrokeGeometry: true });
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 200, y: 154 }).elementId, 'line');
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 200, y: 156 }), undefined);
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 112, y: 150 }).elementId, 'line');
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 109, y: 150 }), undefined);
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 150, y: 170 }).elementId, 'dashed');
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 220, y: 170 }), undefined);
  assert.ok(Object.isFrozen(items[0].strokeGeometry.points[0]));
  assert.throws(() => hitTestSvgConnectorStrokes(items, { x: NaN, y: 0 }), TypeError);
  const singular = structuredClone(items); singular[0].strokeGeometry.transform = [0, 0, 0, 0, 0, 0];
  assert.throws(() => hitTestSvgConnectorStrokes(singular, { x: 0, y: 0 }), RangeError);
  assert.throws(() => getSvgSlideSelectionGeometry(layout, 0, { includeConnectorStrokeGeometry: 1 }), TypeError);
});

test('combined outline and stroke hits preserve object and owner stacking', () => {
  const doc = createPresentation({ masters: [{ id: 'master', name: 'Master', elements: [shape('master-shape', box(0, 0, 100, 100))] }] });
  doc.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master' });
  doc.addSlide({ id: 'page', layoutId: 'layout', elements: [
    { type: 'connector', id: 'line', start: { x: 0, y: 50 }, end: { x: 100, y: 50 }, style: { paint: { type: 'solid', color: '000000' }, width: 2, dash: 'dash' } },
    { type: 'shape', id: 'ellipse', shape: 'ellipse', box: box(40, 40, 20, 20) },
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const items = getSvgSlideSelectionGeometry(layout, 0, { includeConnectorStrokeGeometry: true });
  assert.equal(hitTestSvgSelectionOutlinesAndStrokes(items, { x: 50, y: 50 }).elementId, 'ellipse');
  assert.equal(hitTestSvgSelectionOutlinesAndStrokes(items, { x: 20, y: 50 }).elementId, 'line');
  assert.deepEqual(hitTestSvgSelectionOutlinesAndStrokes(items, { x: 11, y: 50 }).owner, { type: 'master', id: 'master' });
  assert.equal(hitTestSvgSelectionOutlinesAndStrokes(items, { x: 150, y: 150 }), undefined);
  assert.equal(hitTestSvgSelectionOutlinesAndStrokes(getSvgSlideSelectionGeometry(layout, 0), { x: 20, y: 50 }).elementId, 'master-shape');
  assert.throws(() => hitTestSvgSelectionOutlinesAndStrokes([], { x: NaN, y: 0 }), TypeError);
});

test('filled markers remain selectable with transparent lines and nested scaling', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'group', id: 'group', box: box(100, 100, 200, 100), coordinateSize: { width: 100, height: 100 }, children: [
    { type: 'connector', id: 'line', start: { x: 20, y: 50 }, end: { x: 80, y: 50 }, style: { paint: { type: 'solid', color: '000000', opacity: 0 }, width: 2, beginArrow: 'triangle', endArrow: 'diamond' } },
  ] }] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const items = getSvgSlideSelectionGeometry(layout, 0, { includeConnectorStrokeGeometry: true });
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 150, y: 150 }), undefined);
  assert.equal(hitTestSvgConnectorMarkers(items, { x: 130, y: 150 }).elementId, 'line');
  assert.equal(hitTestSvgConnectorMarkers(items, { x: 250, y: 150 }).elementId, 'line');
  assert.equal(hitTestSvgConnectorMarkers(items, { x: 200, y: 150 }), undefined);
  assert.equal(hitTestSvgSelectionOutlinesAndStrokes(items, { x: 130, y: 150 }).elementId, 'line');
  assert.ok(Object.isFrozen(items[0].markerGeometry.points[0]));
  assert.throws(() => hitTestSvgConnectorMarkers(items, { x: Infinity, y: 0 }), TypeError);
  const singular = structuredClone(items); singular[0].markerGeometry.transform = [0, 0, 0, 0, 0, 0];
  assert.throws(() => hitTestSvgConnectorMarkers(singular, { x: 0, y: 0 }), RangeError);
});

test('paint-aware selection passes through unfilled shape interiors and preserves conservative fallbacks', () => {
  const none = { type: 'none' }, ink = { type: 'solid', color: '000000' };
  const doc = createPresentation();
  doc.addSlide({ elements: [
    shape('back', box(0, 0, 100, 100), { style: { fill: ink } }),
    shape('outline', box(20, 20, 60, 60), { style: { fill: none, stroke: { paint: ink, width: 4 } } }),
    shape('clear', box(0, 0, 100, 100), { style: { fill: { type: 'solid', color: 'FF0000', opacity: 0 }, stroke: { paint: none } } }),
    shape('clear-gradient', box(0, 0, 100, 100), { style: { fill: { type: 'linearGradient', angle: 0, stops: [{ position: 0, color: 'FF0000', opacity: 0 }, { position: 1, color: '0000FF', opacity: 0 }] }, stroke: { paint: none } } }),
    shape('pattern', box(300, 0, 60, 60), { style: { fill: { type: 'pattern', pattern: 'cross', foreground: 'FF0000', background: '0000FF', foregroundOpacity: 1, backgroundOpacity: 0 }, stroke: { paint: none } } }),
    shape('text', box(120, 0, 60, 60), { style: { fill: none, stroke: { paint: none } }, text: { content: 'Text' } }),
    shape('wide', box(200, 0, 60, 30), { shape: 'roundRect', style: { fill: none, stroke: { paint: ink, width: 40 } } }),
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const items = getSvgSlideSelectionGeometry(layout, 0, { includeShapePaintSummary: true, includeShapeStrokeGeometry: true });
  assert.equal(hitTestSvgSelectionPaintAware(items, { x: 50, y: 50 }).elementId, 'back');
  assert.equal(hitTestSvgSelectionPaintAware(items, { x: 19, y: 50 }).elementId, 'outline');
  assert.equal(hitTestSvgSelectionPaintAware(items, { x: 140, y: 30 }).elementId, 'text');
  assert.equal(hitTestSvgSelectionPaintAware(items, { x: 230, y: 15 }).elementId, 'wide');
  assert.equal(hitTestSvgSelectionPaintAware(items, { x: 330, y: 30 }).elementId, 'pattern');
  assert.equal(hitTestSvgSelectionPaintAware(getSvgSlideSelectionGeometry(layout, 0), { x: 50, y: 50 }).elementId, 'clear-gradient');
  assert.ok(Object.isFrozen(items[1].shapePaint));
  assert.throws(() => getSvgSlideSelectionGeometry(layout, 0, { includeShapePaintSummary: 1 }), TypeError);
  assert.throws(() => hitTestSvgSelectionPaintAware([], { x: NaN, y: 0 }), TypeError);
});

test('rectangular strokes include outer edges and corners through group scaling', () => {
  const doc = createPresentation();
  const stroke = { paint: { type: 'solid', color: '000000' }, width: 4 };
  doc.addSlide({ elements: [{ type: 'group', id: 'group', box: box(100, 100, 200, 100), coordinateSize: { width: 100, height: 100 }, children: [shape('rect', box(20, 20, 40, 40), { style: { fill: { type: 'none' }, stroke } })] }] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const plain = getSvgSlideSelectionGeometry(layout, 0), items = getSvgSlideSelectionGeometry(layout, 0, { includeShapeStrokeGeometry: true });
  assert.equal(hitTestSvgShapeStrokes(plain, { x: 139, y: 140 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(items, { x: 137, y: 140 }).elementId, 'rect');
  assert.equal(hitTestSvgShapeStrokes(items, { x: 135, y: 140 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(items, { x: 180, y: 140 }), undefined);
  assert.equal(hitTestSvgSelectionOutlinesAndStrokes(items, { x: 137, y: 140 }).elementId, 'rect');
  assert.ok(Object.isFrozen(items[0].shapeStrokeGeometry.points[0]));
  assert.throws(() => getSvgSlideSelectionGeometry(layout, 0, { includeShapeStrokeGeometry: 'true' }), TypeError);
  for (const style of [{ ...stroke, width: 0 }, { ...stroke, paint: { type: 'none' } }, { ...stroke, paint: { type: 'solid', color: '000000', opacity: 0 } }]) {
    const candidate = createPresentation(); candidate.addSlide({ elements: [shape('unsupported', box(0, 0, 40, 40), { style: { fill: { type: 'none' }, stroke: style } })] });
    const other = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(candidate)), 0, { includeShapeStrokeGeometry: true });
    assert.equal(hitTestSvgShapeStrokes(other, { x: 0, y: 20 }), undefined);
  }
});

test('dashed rectangular strokes preserve gaps and phase around corners', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [shape('rect', box(0, 0, 20, 20), { style: { fill: { type: 'none' }, stroke: { paint: { type: 'solid', color: '000000' }, width: 2, dash: 'dash' } } })] });
  const items = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(doc)), 0, { includeShapeStrokeGeometry: true });
  assert.equal(hitTestSvgShapeStrokes(items, { x: 4, y: 0 }).elementId, 'rect');
  assert.equal(hitTestSvgShapeStrokes(items, { x: 11, y: 0 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(items, { x: 20, y: 0 }).elementId, 'rect');
  assert.deepEqual(items[0].shapeStrokeGeometry.dash, [8, 6]);
  assert.ok(Object.isFrozen(items[0].shapeStrokeGeometry.dash));
});

test('polygonal and rounded shape strokes use actual preset edges', () => {
  const doc = createPresentation();
  const style = { fill: { type: 'none' }, stroke: { paint: { type: 'solid', color: '000000' }, width: 4 } };
  doc.addSlide({ elements: [shape('triangle', box(0, 0, 100, 100), { shape: 'triangle', style }), shape('ellipse', box(200, 0, 100, 100), { shape: 'ellipse', style }), shape('round', box(400, 0, 100, 100), { shape: 'roundRect', style })] });
  const items = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(doc)), 0, { includeShapeStrokeGeometry: true });
  assert.equal(hitTestSvgShapeStrokes(items, { x: 50, y: -1 }).elementId, 'triangle');
  assert.equal(hitTestSvgShapeStrokes(items, { x: 50, y: 50 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(items, { x: 0, y: 0 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(items, { x: 250, y: 0 }).elementId, 'ellipse');
  assert.equal(hitTestSvgShapeStrokes(items, { x: 450, y: 0 }).elementId, 'round');
  assert.equal(hitTestSvgSelectionOutlinesAndStrokes(items, { x: 50, y: -1 }).elementId, 'triangle');
});

test('ellipse stroke uses distance to its curve rather than radial scaling', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [shape('ellipse', box(0, 0, 100, 20), { shape: 'ellipse', style: { fill: { type: 'none' }, stroke: { paint: { type: 'solid', color: '000000' }, width: 4 } } })] });
  const items = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(doc)), 0, { includeShapeStrokeGeometry: true });
  assert.equal(hitTestSvgShapeStrokes(items, { x: 50, y: -1 }).elementId, 'ellipse');
  assert.equal(hitTestSvgShapeStrokes(items, { x: 50, y: 10 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(items, { x: 100.5, y: 10 }).elementId, 'ellipse');
  assert.equal(hitTestSvgShapeStrokes(items, { x: 50, y: -3 }), undefined);
  const wide = createPresentation();
  wide.addSlide({ elements: [shape('wide', box(0, 0, 100, 20), { shape: 'ellipse', style: { fill: { type: 'none' }, stroke: { paint: { type: 'solid', color: '000000' }, width: 19.8 } } })] });
  const wideItems = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(wide)), 0, { includeShapeStrokeGeometry: true });
  assert.equal(hitTestSvgShapeStrokes(wideItems, { x: 60, y: 10 }).elementId, 'wide');
  assert.equal(hitTestSvgShapeStrokes(wideItems, { x: 60, y: 10 + 1e-14 }).elementId, 'wide');
});

test('rounded rectangle stroke handles interior, exterior and wide strokes', () => {
  const make = width => {
    const doc = createPresentation();
    doc.addSlide({ elements: [shape('round', box(0, 0, 100, 50), { shape: 'roundRect', style: { fill: { type: 'none' }, stroke: { paint: { type: 'solid', color: '000000' }, width } } })] });
    return getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(doc)), 0, { includeShapeStrokeGeometry: true });
  };
  const narrow = make(4), wide = make(60);
  assert.equal(hitTestSvgShapeStrokes(narrow, { x: 50, y: -1 }).elementId, 'round');
  assert.equal(hitTestSvgShapeStrokes(narrow, { x: 50, y: 25 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(narrow, { x: -2, y: -2 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(wide, { x: 50, y: 25 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(wide, { x: -15, y: 25 }), undefined);
  assert.equal(hitTestSvgShapeStrokes(make(50), { x: 50, y: 0 }), undefined);
  assert.ok(Object.isFrozen(narrow[0].shapeStrokeGeometry.box));
});

test('zero-length connector markers use SVG default tangent and remain selectable', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'connector', id: 'zero', start: { x: 50, y: 50 }, end: { x: 50, y: 50 }, style: { paint: { type: 'solid', color: '000000', opacity: 0 }, width: 2, beginArrow: 'triangle', endArrow: 'triangle' } }] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc));
  const items = getSvgSlideSelectionGeometry(layout, 0, { includeConnectorStrokeGeometry: true });
  assert.equal(hitTestSvgConnectorMarkers(items, { x: 42, y: 50 }).elementId, 'zero');
  assert.equal(hitTestSvgSelectionOutlinesAndStrokes(items, { x: 42, y: 50 }).elementId, 'zero');
  assert.equal(hitTestSvgConnectorMarkers(items, { x: 55, y: 50 }), undefined);
});

test('open arrow marker uses a hollow, clipped mitered stroke', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'connector', id: 'arrow', start: { x: 20, y: 50 }, end: { x: 80, y: 50 }, style: { paint: { type: 'solid', color: '000000', opacity: 0 }, width: 2, endArrow: 'arrow' } }] });
  const items = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(doc)), 0, { includeConnectorStrokeGeometry: true });
  assert.equal(hitTestSvgConnectorMarkers(items, { x: 72, y: 46 }).elementId, 'arrow');
  assert.equal(hitTestSvgConnectorMarkers(items, { x: 72, y: 50 }), undefined);
  assert.equal(hitTestSvgConnectorMarkers(items, { x: 82.2, y: 50 }), undefined);
  assert.equal(hitTestSvgSelectionOutlinesAndStrokes(items, { x: 72, y: 46 }).elementId, 'arrow');
});

test('dash phase continues through bends and round caps extend into gaps', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'connector', id: 'line', start: { x: 0, y: 0 }, route: [{ x: 5, y: 0 }], end: { x: 5, y: 30 }, style: { paint: { type: 'solid', color: '000000' }, width: 2, dash: 'dash' } }] });
  const items = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(doc)), 0, { includeConnectorStrokeGeometry: true });
  const hit = (x, y) => !!hitTestSvgConnectorStrokes(items, { x, y });
  assert.equal(hit(5, 3.9), true);
  assert.equal(hit(5, 4.1), false);
  assert.equal(hit(5, 8.1), true);
  assert.equal(hit(5, 10), true);
  assert.deepEqual(items[0].strokeGeometry.dash, [8, 6]);
  assert.ok(Object.isFrozen(items[0].strokeGeometry.dash));
});

test('captured stroke geometry covers closed seams and zero-length strokes', () => {
  const doc = createPresentation();
  const style = { paint: { type: 'solid', color: '000000' }, width: 4 };
  doc.addSlide({ elements: [
    { type: 'connector', id: 'closed', start: { x: 0, y: 0 }, route: [{ x: 20, y: 0 }, { x: 20, y: 20 }], end: { x: 0, y: 0 }, pathMode: 'closed-round', style },
    { type: 'connector', id: 'zero', start: { x: 50, y: 50 }, end: { x: 50, y: 50 }, style },
  ] });
  const items = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(doc)), 0, { includeConnectorStrokeGeometry: true });
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 10, y: 10 }).elementId, 'closed');
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 51, y: 50 }).elementId, 'zero');
  assert.equal(hitTestSvgConnectorStrokes(items, { x: 52.1, y: 50 }), undefined);
});

test('zero-length connector segments remain selectable within the explicit tolerance', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'connector', id: 'point', start: { x: 10, y: 20 }, end: { x: 10, y: 20 } }] });
  const items = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(doc)), 0);
  assert.equal(hitTestSvgSelectionGeometry(items, { x: 13, y: 24 }, 5).elementId, 'point');
  assert.equal(hitTestSvgSelectionGeometry(items, { x: 13, y: 24 }, 4), undefined);
  const huge = structuredClone(items); huge[0].points[0].x = -Number.MAX_VALUE;
  assert.throws(() => hitTestSvgSelectionGeometry(huge, { x: Number.MAX_VALUE, y: 0 }, 1), RangeError);
  const malformed = resolveNormalizedLayout(normalizePresentation(doc)); malformed.slides[0].elements[0].route = new Array(1000000);
  assert.throws(() => getSvgSlideSelectionGeometry(malformed, 0), RangeError);
});

test('outline hits distinguish preset interior from frame and preserve stacking', () => {
  const cases = [
    ['rect', { x: 1, y: 1 }, true], ['ellipse', { x: 1, y: 1 }, false], ['roundRect', { x: 0, y: 0 }, false],
    ['triangle', { x: 5, y: 5 }, false], ['diamond', { x: 5, y: 5 }, false],
    ['arrow', { x: 5, y: 5 }, false], ['chevron', { x: 5, y: 50 }, false], ['parallelogram', { x: 2, y: 2 }, false],
  ];
  for (const [kind, point, expected] of cases) {
    const doc = createPresentation();
    doc.addSlide({ elements: [shape('under', box(0, 0, 100, 100)), { type: 'shape', id: 'top', shape: kind, box: box(0, 0, 100, 100) }] });
    const items = getSvgSlideSelectionGeometry(resolveNormalizedLayout(normalizePresentation(doc)), 0);
    assert.equal(hitTestSvgSelectionGeometry(items, point, 0).elementId, 'top', kind);
    assert.equal(hitTestSvgSelectionOutlines(items, point, 0).elementId, expected ? 'top' : 'under', kind);
    assert.equal(hitTestSvgSelectionOutlines(items, { x: 70, y: 50 }, 0).elementId, 'top', kind);
    if (items[1].outline) assert.ok(Object.isFrozen(items[1].outline));
  }
});
