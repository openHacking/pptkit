import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveNormalizedLayout, resolveShapeTextBox } from '@pptkit/layout';
import { renderIncrementalSvg, renderLayoutToSvg } from '../dist/index.js';

const box = { x: 0, y: 0, width: 300, height: 100 };
const paragraph = runs => [{ runs }];
const mixed = paragraph([{ text: 'Alpha ', style: { color: 'FF0000', fontSize: 18 } }, { text: 'Beta', style: { bold: true, underline: true, fontSize: 18 } }]);
function layout(elements) {
  const d = createPresentation(); d.addSlide({ id: 's', elements });
  return resolveNormalizedLayout(normalizePresentation(d));
}
const item = (id, content = mixed, frame = { wrap: false }) => ({ type: 'text', id, box, content, frame });

test('opt-in native no-wrap mixed runs preserve text and styling without changing default', async () => {
  const doc = layout([item('mixed')]);
  const original = await renderLayoutToSvg(doc);
  const explicit = await renderLayoutToSvg(doc, { mixedTextMode: 'browser' });
  assert.deepEqual(explicit, original);
  assert.match(original.slides[0].svg, /<foreignObject/);
  const result = await renderLayoutToSvg(doc, { mixedTextMode: 'native-no-wrap', portabilityDiagnostics: 'warn' });
  const svg = result.slides[0].svg;
  assert.doesNotMatch(svg, /<foreignObject/);
  assert.match(svg, /fill="#FF0000"[^>]*>Alpha <\/tspan>/);
  assert.match(svg, /font-weight="700"[^>]*text-decoration="underline"[^>]*>Beta<\/tspan>/);
  assert.deepEqual(result.warnings.map(w => w.code), ['svg-native-mixed-text-approximation', 'svg-system-font-dependency']);
  await assert.rejects(renderLayoutToSvg(doc, { mixedTextMode: null }), TypeError);
});

test('wrapping, bidi, tabs and run actions retain the XHTML fallback', async () => {
  const linked = paragraph([{ text: 'A', action: { type: 'url', url: 'https://example.invalid' } }, { text: 'B' }]);
  const bidi = paragraph([{ text: 'A' }, { text: 'שלום' }]);
  const tab = paragraph([{ text: 'A\t' }, { text: 'B' }]);
  const doc = layout([item('wrap', mixed, { wrap: true }), item('bidi', bidi), item('tab', tab), item('link', linked)]);
  const result = await renderLayoutToSvg(doc, { mixedTextMode: 'native-no-wrap', portabilityDiagnostics: 'warn' });
  assert.deepEqual(result.warnings.filter(w => w.code === 'svg-foreign-object-dependency').map(w => w.elementId), ['wrap', 'bidi', 'tab', 'link']);
  assert.ok(!result.warnings.some(w => w.code === 'svg-native-mixed-text-approximation'));
  assert.match(result.slides[0].svg, /href="https:\/\/example.invalid"/);
});

test('native mixed text keeps hard breaks, whitespace, escaping and combining clusters', async () => {
  const content = paragraph([{ text: ' A&', style: { color: 'FF0000' } }, { text: '<e', style: { bold: true } }, { text: '\u0301\r\n B ' }]);
  const result = await renderLayoutToSvg(layout([item('special', content)]), { mixedTextMode: 'native-no-wrap' });
  const svg = result.slides[0].svg;
  assert.equal((svg.match(/data-pptkit-text-line=/g) ?? []).length, 2);
  assert.match(svg, /> A&amp;<\/tspan>/);
  assert.match(svg, />&lt;e<\/tspan>/);
  assert.match(svg, />\u0301<\/tspan>/);
  assert.match(svg, /> B <\/tspan>/);
  assert.match(svg, /<tspan style="display:none">&#13;\n<\/tspan>/);
});

test('native mode admits wrapped mixed text and single-run shape text', async () => {
  const content = paragraph([{ text: 'Alpha ', style: { fontSize: 18 } }, { text: 'Beta Gamma', style: { fontSize: 18, bold: true } }]);
  const doc = layout([
    { ...item('wrapped', content, { wrap: true }), box: { x: 0, y: 0, width: 65, height: 120 } },
    { type: 'shape', id: 'shape', shape: 'rect', box: { x: 100, y: 0, width: 65, height: 120 }, text: { content: 'One two three', frame: { wrap: true } } },
  ]);
  const before = await renderLayoutToSvg(doc, { mixedTextMode: 'native-no-wrap' });
  assert.equal((before.slides[0].svg.match(/<foreignObject/g) ?? []).length, 2);
  const result = await renderLayoutToSvg(doc, { mixedTextMode: 'native', portabilityDiagnostics: 'warn' });
  assert.doesNotMatch(result.slides[0].svg, /<foreignObject/);
  assert.ok((result.slides[0].svg.match(/data-pptkit-text-line=/g) ?? []).length >= 4);
  assert.match(result.slides[0].svg, /font-size="0"> <\/tspan>/);
  assert.deepEqual(result.warnings.filter(w => w.code === 'svg-native-mixed-text-approximation').map(w => w.elementId), ['wrapped', 'shape']);
  assert.equal(result.warnings.filter(w => w.code === 'svg-system-font-dependency').length, 2);
  const initial = await renderIncrementalSvg(doc, { environmentRevision: 'v', mixedTextMode: 'native-no-wrap' });
  const changed = await renderIncrementalSvg(doc, { environmentRevision: 'v', mixedTextMode: 'native' }, initial.cache);
  assert.deepEqual(changed.renderedSlideIds, ['s']);
  assert.deepEqual(changed.result, await renderLayoutToSvg(doc, { mixedTextMode: 'native' }));
});

test('native mode retains fallback for direction-sensitive text and run links', async () => {
  const d = layout([item('bidi', paragraph([{ text: 'A' }, { text: 'שלום' }]), { wrap: true }), item('link', paragraph([{ text: 'A', action: { type: 'url', url: 'https://example.invalid' } }, { text: 'B' }]), { wrap: true })]);
  const result = await renderLayoutToSvg(d, { mixedTextMode: 'native', portabilityDiagnostics: 'warn' });
  assert.deepEqual(result.warnings.filter(w => w.code === 'svg-foreign-object-dependency').map(w => w.elementId), ['bidi', 'link']);
  assert.ok(!result.warnings.some(w => w.code === 'svg-native-mixed-text-approximation'));
});

test('opt-in shape text uses resolved text geometry and preserves the outer transform', async () => {
  const d = layout([{ type: 'shape', id: 'shape', shape: 'ellipse', box: { x: 20, y: 30, width: 240, height: 120 }, transform: { rotation: 15, flipH: true }, text: { content: mixed, frame: { wrap: false, margin: 0 } } }]);
  const shape = d.slides[0].elements[0];
  const textBox = resolveShapeTextBox(shape);
  const normal = await renderLayoutToSvg(d);
  assert.deepEqual(normal, await renderLayoutToSvg(d, { mixedTextMode: 'browser' }));
  const selected = await renderLayoutToSvg(d, { mixedTextMode: 'native-no-wrap', portabilityDiagnostics: 'warn' });
  assert.match(normal.slides[0].svg, /<foreignObject/);
  assert.doesNotMatch(selected.slides[0].svg, /<foreignObject/);
  assert.match(selected.slides[0].svg, /transform="[^"]*rotate\(15\)/);
  assert.match(selected.slides[0].svg, /scale\(-1 1\)/);
  assert.match(selected.slides[0].svg, new RegExp(`<text x="${textBox.x}"`));
  assert.match(selected.slides[0].svg, />Alpha <\/tspan>/);
  assert.match(selected.slides[0].svg, />Beta<\/tspan>/);
  assert.deepEqual(selected.warnings.map(w => [w.code, w.elementId]), [['svg-native-mixed-text-approximation', 'shape'], ['svg-system-font-dependency', 'shape']]);
  const cached = await renderIncrementalSvg(d, { environmentRevision: 'v' });
  const changed = await renderIncrementalSvg(d, { environmentRevision: 'v', mixedTextMode: 'native-no-wrap', portabilityDiagnostics: 'warn' }, cached.cache);
  assert.deepEqual(changed.renderedSlideIds, ['s']);
  assert.deepEqual(changed.result, selected);
});

test('shape text with wrapping, bidi, links or vertical writing stays in XHTML', async () => {
  const shape = (id, content, frame) => ({ type: 'shape', id, shape: 'rect', box, text: { content, frame } });
  const d = layout([
    shape('wrap', mixed, { wrap: true }),
    shape('bidi', paragraph([{ text: 'A' }, { text: 'שלום' }]), { wrap: false }),
    shape('link', paragraph([{ text: 'A', action: { type: 'url', url: 'https://example.invalid' } }, { text: 'B' }]), { wrap: false }),
    shape('vertical', mixed, { wrap: false, textDirection: 'rotate90' }),
  ]);
  const result = await renderLayoutToSvg(d, { mixedTextMode: 'native-no-wrap', portabilityDiagnostics: 'warn' });
  assert.deepEqual(result.warnings.filter(w => w.code === 'svg-foreign-object-dependency').map(w => w.elementId), ['wrap', 'bidi', 'link', 'vertical']);
  assert.ok(!result.warnings.some(w => w.code === 'svg-native-mixed-text-approximation'));
});

test('incremental mode changes invalidate the cache and agree with full rendering', async () => {
  const doc = layout([item('mixed')]);
  const first = await renderIncrementalSvg(doc, { environmentRevision: 'v' });
  const options = { environmentRevision: 'v', mixedTextMode: 'native-no-wrap' };
  const next = await renderIncrementalSvg(doc, options, first.cache);
  assert.deepEqual(next.renderedSlideIds, ['s']);
  assert.deepEqual(next.result, await renderLayoutToSvg(doc, options));
  assert.deepEqual((await renderIncrementalSvg(doc, options, next.cache)).renderedSlideIds, []);
  await assert.rejects(renderIncrementalSvg(doc, { environmentRevision: 'v', mixedTextMode: null }), TypeError);
});
