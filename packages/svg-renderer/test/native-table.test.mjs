import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveNormalizedLayout } from '@pptkit/layout';
import { renderIncrementalSvg, renderLayoutToSvg } from '../dist/index.js';

function fixture() {
  const d = createPresentation();
  d.addSlide({ id: 's', elements: [{ type: 'table', id: 'table', box: { x: 20, y: 30, width: 180, height: 80 }, transform: { rotation: 10 }, columns: [60, 60, 60], rows: [
    { height: 40, cells: [{ content: 'Head', colSpan: 2, headerScope: 'column', style: { fill: { type: 'solid', color: 'FF0000', opacity: 0.5 } } }, { content: 'Right', rowSpan: 2 }] },
    { height: 40, cells: [{ content: 'Left' }, { content: [{ runs: [{ text: 'Mix ' }, { text: 'bold', style: { bold: true } }] }] }] },
  ] }] });
  return resolveNormalizedLayout(normalizePresentation(d));
}

test('native tables use resolved merge geometry, clipping, paint and accessible grid roles', async () => {
  const layout = fixture(), plain = await renderLayoutToSvg(layout);
  assert.deepEqual(plain, await renderLayoutToSvg(layout, { tableMode: 'browser' }));
  assert.match(plain.slides[0].svg, /<foreignObject/);
  const result = await renderLayoutToSvg(layout, { tableMode: 'native', portabilityDiagnostics: 'warn' });
  const svg = result.slides[0].svg;
  assert.doesNotMatch(svg, /<foreignObject/);
  assert.match(svg, /role="table"[^>]*aria-rowcount="2" aria-colcount="3"/);
  assert.equal((svg.match(/role="row"/g) ?? []).length, 2);
  assert.match(svg, /role="columnheader" aria-colindex="1" aria-rowindex="1" aria-colspan="2" aria-rowspan="1"/);
  assert.match(svg, /aria-colindex="3" aria-rowindex="1" aria-colspan="1" aria-rowspan="2"/);
  assert.match(svg, /aria-label="Mix bold" data-pptkit-cell="1:1"/);
  assert.match(svg, /<rect x="20" y="30" width="120" height="40" fill="#FF0000" fill-opacity="0.5"/);
  assert.match(svg, /<rect x="140" y="30" width="60" height="80"/);
  assert.match(svg, /transform="[^"]*rotate\(10\)/);
  assert.equal((svg.match(/data-pptkit-cell=/g) ?? []).length, 4);
  assert.equal((svg.match(/<clipPath /g) ?? []).length, 5);
  assert.match(svg, />Mix<\/tspan><tspan font-size="0"> <\/tspan>/); assert.match(svg, /font-weight="700"[^>]*>bold<\/tspan>/);
  assert.ok(result.warnings.some(w => w.code === 'svg-native-table-approximation'));
  assert.ok(result.warnings.some(w => w.code === 'svg-system-font-dependency'));
  assert.ok(!result.warnings.some(w => w.code === 'svg-foreign-object-dependency'));
});

test('native cell names escape markup and preserve paragraph and list text', async () => {
  const d = createPresentation();
  d.addSlide({ id: 's', elements: [{ type: 'table', id: 't', box: { x: 0, y: 0, width: 100, height: 40 }, columns: [100], rows: [{ height: 40, cells: [{ content: [
    { style: { bullet: { type: 'number', startAt: 3 } }, runs: [{ text: 'A & ' }, { text: '<B>' }] },
    { runs: [{ text: 'Next' }] },
  ] }] }] }] });
  const result = await renderLayoutToSvg(resolveNormalizedLayout(normalizePresentation(d)), { tableMode: 'native' });
  assert.match(result.slides[0].svg, /aria-label="3\. A &amp; &lt;B&gt;&#10;Next"/);
});

test('unsupported directions and run actions fall back atomically', async () => {
  const source = fixture();
  for (const mutate of [
    cell => { cell.style.textDirection = 'eastAsianVertical'; },
    cell => { cell.content[0].runs[0].text = 'שלום'; },
    cell => { cell.content[0].runs[0].text = 'a\tb'; },
    cell => { cell.content[0].runs[0].action = { type: 'url', url: 'https://example.invalid' }; },
    cell => { cell.colSpan = 99; },
  ]) {
    const layout = structuredClone(source); mutate(layout.slides[0].elements[0].rows[0].cells[0]);
    const result = await renderLayoutToSvg(layout, { tableMode: 'native', portabilityDiagnostics: 'warn' });
    assert.ok(result.warnings.some(w => w.code === 'svg-native-table-fallback'));
    assert.ok(result.warnings.some(w => w.code === 'svg-foreign-object-dependency'));
    assert.match(result.slides[0].svg, /<foreignObject/);
    assert.ok(!result.warnings.some(w => w.code === 'svg-native-table-approximation'));
  }
});

test('quarter-turn cell text rotates inside physical clips with mapped margins', async () => {
  const d = createPresentation();
  const margin = { top: 3, right: 5, bottom: 7, left: 11 };
  d.addSlide({ id: 's', elements: [{ type: 'table', id: 't', box: { x: 20, y: 30, width: 200, height: 160 }, columns: [100, 100], rows: [{ height: 160, cells: [
    { content: 'Clockwise', style: { textDirection: 'rotate90', margin } },
    { content: 'Counter', style: { textDirection: 'rotate270', margin } },
  ] }] }] });
  const layout = resolveNormalizedLayout(normalizePresentation(d));
  const normal = await renderLayoutToSvg(layout), result = await renderLayoutToSvg(layout, { tableMode: 'native', portabilityDiagnostics: 'warn' });
  assert.match(normal.slides[0].svg, /<foreignObject/);
  assert.doesNotMatch(result.slides[0].svg, /<foreignObject/);
  assert.match(result.slides[0].svg, /transform="rotate\(90 70 110\)"><text x="-7" y="79\.4"/);
  assert.match(result.slides[0].svg, /transform="rotate\(270 170 110\)"><text x="97" y="85\.4"/);
  assert.match(result.slides[0].svg, /<clipPath[^>]+><rect x="20" y="30" width="100" height="160"/);
  assert.equal(result.warnings.filter(w => w.code === 'table-text-direction-approximation').length, 1);
  assert.ok(!result.warnings.some(w => w.code === 'svg-native-table-fallback'));
});

test('native tables use one resolved shared border, preserving width and opacity', async () => {
  const d = createPresentation();
  const border = (color, width, opacity) => ({ paint: { type: 'solid', color, opacity }, width });
  d.addSlide({ id: 's', elements: [{ type: 'table', id: 't', box: { x: 0, y: 0, width: 100, height: 40 }, columns: [50, 50], rows: [{ height: 40, cells: [
    { content: 'A', style: { borders: { right: border('FF0000', 2, 1) } } },
    { content: 'B', style: { borders: { left: border('00FF00', 3, 0.5) } } },
  ] }] }] });
  const layout = resolveNormalizedLayout(normalizePresentation(d));
  const svg = (await renderLayoutToSvg(layout, { tableMode: 'native' })).slides[0].svg;
  const boundary = svg.match(/<line\b[^>]*data-pptkit-table-border="v:0:1"\/>/g);
  assert.equal(boundary?.length, 1);
  assert.match(boundary[0], /stroke="#00FF00" stroke-opacity="0.5" stroke-width="3"/);
  assert.match(boundary[0], /x1="50" y1="0" x2="50" y2="40"/);
});

test('incremental cache respects native table mode and invalid values reject', async () => {
  const layout = fixture(), first = await renderIncrementalSvg(layout, { environmentRevision: 'v' });
  const options = { environmentRevision: 'v', tableMode: 'native' };
  const next = await renderIncrementalSvg(layout, options, first.cache);
  assert.deepEqual(next.renderedSlideIds, ['s']);
  assert.deepEqual(next.result, await renderLayoutToSvg(layout, options));
  assert.deepEqual((await renderIncrementalSvg(layout, options, next.cache)).renderedSlideIds, []);
  await assert.rejects(renderIncrementalSvg(layout, { environmentRevision: 'v', tableMode: null }), TypeError);
  await assert.rejects(renderLayoutToSvg(layout, { tableMode: null }), TypeError);
});
