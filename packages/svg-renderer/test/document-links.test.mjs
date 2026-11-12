import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '@pptkit/core';
import { resolveLayout } from '@pptkit/layout';
import { renderIncrementalSvg, renderLayoutToSvg } from '../dist/index.js';

function fixture() {
  const deck = createPresentation();
  deck.addSlide({ id: 'target' });
  deck.addSlide({ id: 'source', elements: [
    { type: 'shape', id: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 50, height: 40 }, action: { type: 'slide', slideId: 'target', tooltip: 'Go <there>' } },
    { type: 'text', id: 'text', box: { x: 0, y: 50, width: 120, height: 40 }, content: [{ runs: [{ text: 'Go', action: { type: 'slide', slideId: 'target', tooltip: 'Next' } }] }] },
  ] });
  return resolveLayout(deck);
}

test('document links bind selected slides while default and missing targets remain metadata', async () => {
  const layout = fixture();
  const ordinary = await renderLayoutToSvg(layout);
  assert.equal(ordinary.warnings.filter(w => w.code === 'slide-action-not-linked').length, 2);
  const result = await renderLayoutToSvg(layout, { slideIndices: [1, 0], slideActionMode: 'document-links', idNamespace: 'a b' });
  assert.equal(result.warnings.filter(w => w.code === 'slide-action-not-linked').length, 0);
  const target = /<svg id="([^"]+)"/.exec(result.slides[1].svg)[1];
  const other = await renderLayoutToSvg(layout, { slideIndices: [1, 0], slideActionMode: 'document-links', idNamespace: 'other' });
  assert.notEqual(target, /<svg id="([^"]+)"/.exec(other.slides[1].svg)[1]);
  assert.equal((result.slides[0].svg.match(new RegExp(`href="#${target}"`, 'g')) ?? []).length, 2);
  assert.match(result.slides[0].svg, /aria-label="Go &lt;there&gt;"/);
  assert.match(result.slides[0].svg, /title="Next"/);
  const absent = await renderLayoutToSvg(layout, { slideIndices: [1], slideActionMode: 'document-links' });
  assert.equal(absent.warnings.filter(w => w.code === 'slide-action-not-linked').length, 2);
  assert.doesNotMatch(absent.slides[0].svg, /href="#pptkit-link-slide-/);
  await assert.rejects(renderLayoutToSvg(layout, { slideActionMode: 'bad' }), TypeError);
  const duplicate = { ...layout, slides: [layout.slides[0], { ...layout.slides[1], id: 'target' }] };
  await assert.rejects(renderLayoutToSvg(duplicate, { slideActionMode: 'document-links' }), /unique/);
});

test('incremental document links track all targets and invalidate after page reorder', async () => {
  const layout = fixture(), options = { environmentRevision: '1', slideActionMode: 'document-links' };
  const initial = await renderIncrementalSvg(layout, options);
  assert.deepEqual(initial.result.slides, (await renderLayoutToSvg(layout, options)).slides);
  assert.deepEqual((await renderIncrementalSvg(layout, options, initial.cache)).renderedSlideIds, []);
  const metadata = await renderIncrementalSvg(layout, { environmentRevision: '1' }, initial.cache);
  assert.equal(metadata.renderedSlideIds.length, 2);
  assert.equal(metadata.result.warnings.filter(w => w.code === 'slide-action-not-linked').length, 2);
  const reordered = { ...layout, slides: [...layout.slides].reverse() };
  const changed = await renderIncrementalSvg(reordered, options, initial.cache);
  assert.deepEqual(changed.result.slides, (await renderLayoutToSvg(reordered, options)).slides);
  assert.equal(changed.renderedSlideIds.length, 2);
  await assert.rejects(renderIncrementalSvg(layout, { environmentRevision: '1', slideActionMode: 'bad' }), TypeError);
});

test('decorative slide links do not enter sequential keyboard focus', async () => {
  const deck = createPresentation();
  deck.addSlide({ id: 'target' });
  deck.addSlide({ id: 'source', elements: [
    { type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 50, height: 40 }, accessibility: { decorative: true }, action: { type: 'slide', slideId: 'target' } },
    { type: 'text', box: { x: 0, y: 50, width: 120, height: 40 }, accessibility: { decorative: true }, content: [{ runs: [{ text: 'Go', action: { type: 'slide', slideId: 'target' } }] }] },
  ] });
  const result = await renderLayoutToSvg(resolveLayout(deck), { slideActionMode: 'document-links' });
  assert.equal((result.slides[1].svg.match(/<a href="#pptkit-link-slide-0"[^>]*tabindex="-1"/g) ?? []).length, 2);
});

test('hostile slide IDs cannot inject markup into document links', async () => {
  const deck = createPresentation(), id = 'x"><script>alert(1)</script>';
  deck.addSlide({ id });
  deck.addSlide({ id: 'source', elements: [{ type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 20, height: 20 }, action: { type: 'slide', slideId: id } }] });
  const result = await renderLayoutToSvg(resolveLayout(deck), { slideActionMode: 'document-links' });
  assert.match(result.slides[1].svg, /href="#pptkit-link-slide-0"/);
  assert.match(result.slides[1].svg, /data-pptkit-slide-target="x&quot;&gt;&lt;script&gt;/);
  assert.doesNotMatch(result.slides[1].svg, /<script>/);
});
