import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveNormalizedLayout } from '@pptkit/layout';
import { renderIncrementalSvg, renderLayoutToSvg, SvgRenderAbortError, SvgRenderLimitError } from '../dist/index.js';

function fixture() {
  const doc = createPresentation();
  doc.registerAsset({ id: 'asset', kind: 'image', source: { type: 'url', value: 'https://example.invalid/a.png' }, mimeType: 'image/png', width: 4, height: 4 });
  doc.addSlide({ id: 'a', elements: [{ id: 'image', type: 'image', assetId: 'asset', box: { x: 10, y: 10, width: 40, height: 40 } }] });
  doc.addSlide({ id: 'b', elements: [{ id: 'text', type: 'text', box: { x: 5, y: 5, width: 200, height: 30 }, content: 'hello' }] });
  return resolveNormalizedLayout(normalizePresentation(doc));
}

test('incremental output and warnings match full render across updates, options and moves', async () => {
  const layout = fixture(); let calls = 0;
  const resolveAsset = () => { calls++; return undefined; };
  const options = { environmentRevision: 'v1', resolveAsset, pngAnimationDiagnostics: 'warn', idNamespace: 'cache' };
  const first = await renderIncrementalSvg(layout, options);
  assert.deepEqual(first.renderedSlideIds, ['a', 'b']); assert.equal(calls, 1);
  assert.deepEqual(first.result, await renderLayoutToSvg(layout, options)); calls = 0;
  const unchanged = await renderIncrementalSvg(layout, options, first.cache);
  assert.deepEqual(unchanged.renderedSlideIds, []); assert.equal(calls, 0); assert.deepEqual(unchanged.result, first.result);
  const withAnalysis = structuredClone(layout);
  withAnalysis.analysis = { diagnostics: [{ code: 'current-analysis', severity: 'warning', message: 'current', slideId: 'a' }] };
  const analysisOnly = await renderIncrementalSvg(withAnalysis, options, first.cache);
  assert.deepEqual(analysisOnly.renderedSlideIds, []);
  assert.deepEqual(analysisOnly.result, await renderLayoutToSvg(withAnalysis, options)); calls = 0;
  assert.ok(Object.isFrozen(unchanged.cache.entries[0]) && Object.isFrozen(unchanged.cache.entries[0].warnings[0]));
  unchanged.result.slides[0].svg = 'caller mutation'; unchanged.result.warnings[0].message = 'changed';
  assert.deepEqual((await renderIncrementalSvg(layout, options, unchanged.cache)).result, first.result);
  layout.slides[1].elements[0].box.x += 20;
  const edited = await renderIncrementalSvg(layout, options, first.cache);
  assert.deepEqual(edited.renderedSlideIds, ['b']); assert.equal(calls, 0);
  assert.deepEqual(edited.result, await renderLayoutToSvg(layout, options));
  [layout.slides[0], layout.slides[1]] = [layout.slides[1], layout.slides[0]];
  const moved = await renderIncrementalSvg(layout, options, edited.cache);
  assert.deepEqual(moved.renderedSlideIds, ['b', 'a']);
  assert.deepEqual(moved.result, await renderLayoutToSvg(layout, options));
  const changedMode = await renderIncrementalSvg(layout, { ...options, pngAnimationDiagnostics: 'off' }, moved.cache);
  assert.deepEqual(changedMode.renderedSlideIds, ['b', 'a']);
  const changedRevision = await renderIncrementalSvg(layout, { ...options, environmentRevision: 'v2' }, moved.cache);
  assert.deepEqual(changedRevision.renderedSlideIds, ['b', 'a']);
});

test('cached SVGs still enforce byte limits and cancellation', async () => {
  const layout = fixture(), options = { environmentRevision: 'v1', resolveAsset: () => undefined };
  const first = await renderIncrementalSvg(layout, options);
  const bytes = first.result.slides.reduce((sum, slide) => sum + Buffer.byteLength(slide.svg), 0);
  assert.equal((await renderIncrementalSvg(layout, { ...options, maxSvgBytes: bytes }, first.cache)).renderedSlideIds.length, 0);
  await assert.rejects(renderIncrementalSvg(layout, { ...options, maxSvgBytes: bytes - 1 }, first.cache), SvgRenderLimitError);
  const controller = new AbortController(); controller.abort();
  await assert.rejects(renderIncrementalSvg(layout, { ...options, signal: controller.signal }, first.cache), SvgRenderAbortError);
  await assert.rejects(renderIncrementalSvg(layout, { ...options, slideIndices: [0] }, first.cache), TypeError);
  await assert.rejects(renderIncrementalSvg(layout, { ...options, pngAnimationDiagnostics: null }, first.cache), TypeError);
  await assert.rejects(renderIncrementalSvg(layout, { ...options, resolveAsset: 1 }, first.cache), TypeError);
  assert.deepEqual((await renderIncrementalSvg(layout, options, first.cache)).result, first.result);
});

test('asynchronous input mutation is isolated and resolver mutation is rejected', async () => {
  const layout = fixture(); let release, invoked;
  const entered = new Promise(resolve => { invoked = resolve; });
  const pending = renderIncrementalSvg(layout, { environmentRevision: 'v1', resolveAsset: () => { invoked(); return new Promise(resolve => { release = resolve; }); } });
  await entered; layout.slides[1].elements[0].box.x = 99; release(undefined);
  const first = await pending;
  assert.notEqual(first.result.slides[1].svg, (await renderLayoutToSvg(layout, { resolveAsset: () => undefined })).slides[1].svg);
  const current = await renderIncrementalSvg(layout, { environmentRevision: 'v1', resolveAsset: () => undefined }, first.cache);
  assert.deepEqual(current.renderedSlideIds, ['b']);
  await assert.rejects(renderIncrementalSvg(layout, { environmentRevision: 'v2', resolveAsset: asset => { asset.source.value = 'mutated'; return undefined; } }), /resolver mutated/);
});

test('abort interrupts an in-flight resolver and does not publish a cache', async () => {
  const layout = fixture(), controller = new AbortController(); let called;
  const entered = new Promise(resolve => { called = resolve; });
  const pending = renderIncrementalSvg(layout, { environmentRevision: 'v1', signal: controller.signal, resolveAsset: () => { called(); return new Promise(() => {}); } });
  await entered; controller.abort();
  await assert.rejects(pending, SvgRenderAbortError);
});
