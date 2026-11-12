import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveNormalizedLayout } from '@pptkit/layout';
import { captureSvgPreviewSnapshot, planSvgPreviewInvalidation, renderLayoutToSvg } from '../dist/index.js';

const options = { environmentRevision: 'assets-v1' };
const shape = id => ({ type: 'shape', id, shape: 'rect', box: { x: 10, y: 10, width: 50, height: 50 }, style: { fill: { type: 'solid', color: 'FF0000' } } });
function fixture() {
  const doc = createPresentation({ masters: [{ id: 'm1', name: 'M1', elements: [shape('master')] }, { id: 'm2', name: 'M2' }] });
  doc.defineSlideLayout({ id: 'l1', name: 'L1', masterId: 'm1', elements: [shape('layout')] });
  doc.defineSlideLayout({ id: 'l2', name: 'L2', masterId: 'm2' });
  doc.addSlide({ id: 'a', layoutId: 'l1', elements: [shape('a-shape')] });
  doc.addSlide({ id: 'b', layoutId: 'l1', elements: [shape('b-shape')] });
  doc.addSlide({ id: 'c', layoutId: 'l2', elements: [shape('c-shape')] });
  return resolveNormalizedLayout(normalizePresentation(doc));
}
const ids = items => items.map(item => item.slideId);

test('changed local, shared and global dependencies invalidate the intended pages', async () => {
  const initial = fixture(), before = captureSvgPreviewSnapshot(initial, options);
  const renderedBefore = await renderLayoutToSvg(initial);
  const variants = [
    [layout => { layout.slides[0].elements[0].box.x += 9; }, ['a']],
    [layout => { layout.layouts.find(l => l.id === 'l1').elements[0].box.x += 9; }, ['a', 'b']],
    [layout => { layout.masters.find(m => m.id === 'm1').elements[0].box.x += 9; }, ['a', 'b']],
    [layout => { layout.metadata.title = 'changed'; }, ['a', 'b', 'c']],
    [layout => { layout.assets.push({ id: 'unused', kind: 'image', source: { type: 'url', value: 'https://example.invalid' } }); }, ['a', 'b', 'c']],
    [layout => { [layout.slides[0], layout.slides[1]] = [layout.slides[1], layout.slides[0]]; }, ['b', 'a']],
  ];
  for (const [edit, changed] of variants) {
    const next = structuredClone(initial); edit(next);
    const plan = planSvgPreviewInvalidation(before, captureSvgPreviewSnapshot(next, options));
    assert.deepEqual(ids(plan.render), changed);
    const renderedNext = await renderLayoutToSvg(next);
    for (const item of plan.reusable) assert.equal(renderedNext.slides[item.index].svg, renderedBefore.slides[item.index].svg, item.slideId);
  }
  const newEnv = captureSvgPreviewSnapshot(initial, { environmentRevision: 'assets-v2' });
  assert.deepEqual(ids(planSvgPreviewInvalidation(before, newEnv).render), ['a', 'b', 'c']);
  const newNamespace = captureSvgPreviewSnapshot(initial, { ...options, idNamespace: 'other' });
  assert.deepEqual(ids(planSvgPreviewInvalidation(before, newNamespace).render), ['a', 'b', 'c']);
});

test('snapshots are detached and identify insertions, removals, moves and exact restoration', () => {
  const layout = fixture(), first = captureSvgPreviewSnapshot(layout, options);
  assert.ok(Object.isFrozen(first) && Object.isFrozen(first.slides) && Object.isFrozen(first.slides[0]));
  const unchanged = planSvgPreviewInvalidation(first, captureSvgPreviewSnapshot(layout, options));
  assert.deepEqual(ids(unchanged.reusable), ['a', 'b', 'c']);
  layout.slides[0].elements[0].box.x = 999;
  assert.deepEqual(ids(planSvgPreviewInvalidation(first, captureSvgPreviewSnapshot(layout, options)).render), ['a']);
  layout.slides[0].elements[0].box.x = 10;
  assert.equal(planSvgPreviewInvalidation(first, captureSvgPreviewSnapshot(layout, options)).render.length, 0);
  layout.slides.shift();
  const removed = planSvgPreviewInvalidation(first, captureSvgPreviewSnapshot(layout, options));
  assert.deepEqual(ids(removed.render), ['b', 'c']); assert.deepEqual(removed.removedSlideIds, ['a']);
  const inserted = fixture(); const extra = structuredClone(inserted.slides[2]); extra.id = 'd'; inserted.slides.push(extra);
  const addition = planSvgPreviewInvalidation(first, captureSvgPreviewSnapshot(inserted, options));
  assert.deepEqual(ids(addition.reusable), ['a', 'b', 'c']); assert.deepEqual(ids(addition.render), ['d']);
  assert.throws(() => captureSvgPreviewSnapshot(layout, { environmentRevision: '' }), TypeError);
  assert.throws(() => captureSvgPreviewSnapshot(layout, { ...options, idNamespace: '' }), TypeError);
  const invalid = fixture(); invalid.slides[0].elements[0].box.x = NaN;
  assert.throws(() => captureSvgPreviewSnapshot(invalid, options), TypeError);
  layout.slides[1].id = layout.slides[0].id;
  assert.throws(() => captureSvgPreviewSnapshot(layout, options), TypeError);
  assert.throws(() => planSvgPreviewInvalidation({ schemaVersion: 'unknown' }, first), TypeError);
});
