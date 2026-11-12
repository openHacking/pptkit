import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '../dist/index.js';
import { collectSlideCopyDependencies } from '../dist/document/slide-copy-dependencies.js';
import { setCopiedReference } from '../dist/document/copy-references.js';

function makeSlide() {
  const d = createPresentation({ id: 'source' });
  return d.addSlide({ id: 'page', layoutId: 'layout', customData: { custom: 'opaque' }, notes: [{ runs: [{ text: 'note', action: { type: 'slide', slideId: 'page' } }] }], elements: [
    { id: 'edge', type: 'connector', start: { elementId: 'later' }, end: { elementId: 'external' } },
    { id: 'group', type: 'group', box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [
      { id: 'later', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 }, action: { type: 'url', url: 'https://example.invalid' }, text: { textStylePreset: 'shape-style', content: [{ runs: [{ text: 'x', action: { type: 'slide', slideId: 'other' } }] }] } },
      { id: 'picture', type: 'image', assetId: 'asset', box: { x: 10, y: 10, width: 20, height: 20 } },
    ] },
    { id: 'text', type: 'text', placeholderKey: 'title', textStylePreset: 'body-style', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [{ text: 'stale', field: { id: 'field-id', kind: 'slide-number' } }] }] },
    { id: 'table', type: 'table', box: { x: 0, y: 0, width: 100, height: 30 }, columns: [100], rows: [{ cells: [{ textStylePreset: 'cell-style', content: [{ runs: [{ text: 'cell', action: { type: 'slide', slideId: 'page' } }] }] }] }] },
  ] });
}

test('copy reference writes require existing own paths and cannot touch prototypes', () => {
  const root = { elements: [{ assetId: 'old' }] };
  setCopiedReference(root, 'elements.0.assetId', 'new');
  assert.equal(root.elements[0].assetId, 'new');
  for (const path of ['elements.1.assetId', 'elements.0.absent', 'elements.0.__proto__.polluted', '__proto__.polluted']) {
    assert.throws(() => setCopiedReference(root, path, 'unsafe'), /reference path/);
  }
  assert.equal(Object.prototype.polluted, undefined);
});

test('copy dependency inventory locates nested owners, forward links, text and notes without mutating source', () => {
  const slide = makeSlide(), before = JSON.stringify(slide.elements), result = collectSlideCopyDependencies(slide);
  assert.deepEqual(result.elementIds, ['edge', 'group', 'later', 'picture', 'text', 'table']);
  assert.deepEqual(result.connectors, [{ targetId: 'later', path: 'elements.0.start', slideOwned: true }, { targetId: 'external', path: 'elements.0.end', slideOwned: false }]);
  assert.deepEqual(result.assets, [{ id: 'asset', path: 'elements.1.children.1.assetId' }]);
  assert.deepEqual(result.slideActions.map(a => [a.targetId, a.self]), [['other', false], ['page', true], ['page', true]]);
  assert.deepEqual(result.urlActions, [{ path: 'elements.1.children.0.action' }]);
  assert.deepEqual(result.textPresets.map(p => p.name), ['shape-style', 'body-style', 'cell-style']);
  assert.deepEqual(result.placeholders, [{ key: 'title', path: 'elements.2.placeholderKey' }]);
  assert.deepEqual(result.fields, [{ id: 'field-id', kind: 'slide-number', path: 'elements.2.content.0.runs.0.field' }]);
  assert.equal(result.layoutId, 'layout'); assert.equal(result.hasOpaqueCustomData, true); assert.equal(result.hasTimeline, false);
  assert.ok(Object.isFrozen(result)); assert.ok(Object.isFrozen(result.assets[0])); assert.equal(JSON.stringify(slide.elements), before);
});

test('copy dependency inventory rejects malformed references and enforces work budget without partial output', () => {
  const slide = makeSlide(); assert.throws(() => collectSlideCopyDependencies(slide, 1), /work limit/);
  for (const value of [0, -1, 1.2, Infinity]) assert.throws(() => collectSlideCopyDependencies(slide, value), RangeError);
  const data = { ...slide, id: 'page', elements: [{ id: 'same', type: 'image', assetId: 'a' }, { id: 'same', type: 'shape' }], customData: {} };
  assert.throws(() => collectSlideCopyDependencies(data), /identities/);
  assert.throws(() => collectSlideCopyDependencies({ ...data, elements: [{ id: 'x', type: 'future' }] }), /unsupported element/);
  assert.throws(() => collectSlideCopyDependencies({ ...data, elements: [{ id: 'x', type: 'connector', start: null, end: { x: 1, y: 1 } }] }), /endpoints/);
  assert.throws(() => collectSlideCopyDependencies({ ...data, elements: [{ id: 'x', type: 'text', action: { type: 'macro' }, content: 'x' }] }), /unsupported action/);
  assert.throws(() => collectSlideCopyDependencies({ ...data, elements: Array(1000) }, 2), /work limit/);
  assert.throws(() => collectSlideCopyDependencies({ ...data, elements: [{ id: 'g', type: 'group', children: Array(1000) }] }, 2), /work limit/);
  for (const content of [Array(1), [{ runs: Array(1) }], [{ runs: [null] }]]) {
    assert.throws(() => collectSlideCopyDependencies({ ...data, elements: [{ id: 'x', type: 'text', content }] }), /valid text runs/);
  }
  for (const rows of [Array(1), [{ cells: Array(1) }], [{ cells: [null] }]]) {
    assert.throws(() => collectSlideCopyDependencies({ ...data, elements: [{ id: 'x', type: 'table', rows }] }), /valid table cells/);
  }
});

test('copy dependency inventory validates and identifies owned timeline targets and triggers', () => {
  const d = createPresentation(), slide = d.addSlide({ id: 'p', elements: [{ id: 'target', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 } }, { id: 'button', type: 'shape', shape: 'rect', box: { x: 20, y: 0, width: 10, height: 10 } }] });
  slide.setEventTimeline({ stages: [{ id: 'stage', target: { pageId: 'p', elementId: 'target' }, trigger: { type: 'object-click', elementId: 'button' }, timing: { kind: 'span', id: 'time', durationMs: 10, fill: 'hold' } }] });
  const result = collectSlideCopyDependencies(slide);
  assert.equal(result.hasTimeline, true); assert.deepEqual(result.timelineTargets, [{ stageId: 'stage', targetId: 'target', triggerId: 'button' }]);
  const invalid = { ...slide, id: 'other', elements: slide.elements, eventTimeline: slide.eventTimeline, customData: {} };
  assert.throws(() => collectSlideCopyDependencies(invalid), /target/i);
});
