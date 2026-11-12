import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation, validatePresentation } from '../dist/index.js';
import { multiMasterDocument } from '../../../tests/fixtures/multi-master.mjs';

const box = { x: 10, y: 20, width: 100, height: 50 };
const shape = id => ({ id, type: 'shape', shape: 'rect', box });
const placeholder = () => ({ key: 'same', kind: 'body', box: { ...box } });
const validation = code => error => error.name === 'PresentationValidationError' && error.diagnostics.some(d => d.code === code);

test('master edits propagate to every bound layout while preserving local declarations and old snapshots', () => {
  const d = multiMasterDocument(createPresentation), oldRegistry = d.masters, before = normalizePresentation(d);
  const input = { name: 'New red', theme: { colors: { accent1: '337711' }, fonts: { body: 'Verdana' } }, background: { type: 'none' } };
  const replacement = d.replaceSlideMaster('red', input);
  input.theme.colors.accent1 = 'FFFFFF';
  const n = normalizePresentation(d);
  assert.deepEqual(n.layouts.filter(l => l.masterId === 'red').map(l => [l.theme.colors.accent1, l.theme.fonts.body, l.background.type]), [['337711', 'Verdana', 'none'], ['337711', 'Verdana', 'none']]);
  assert.deepEqual(n.slides[1], before.slides[1]); assert.deepEqual(n.layouts.find(l => l.id === 'layout-1'), before.layouts.find(l => l.id === 'layout-1'));
  assert.deepEqual(n.slides[3], before.slides[3]);
  assert.equal(d.masters[0], replacement); assert.ok(Object.isFrozen(replacement.theme.colors));
  assert.equal(oldRegistry[0].theme.colors.accent1, 'CC2200'); assert.notEqual(oldRegistry, d.masters);
  d.replaceSlideMaster('red', { name: 'Inherit' });
  const inherited = normalizePresentation(d);
  assert.equal(inherited.layouts.find(l => l.id === 'layout-0').theme.colors.accent1, '112233'); assert.equal(inherited.layouts.find(l => l.id === 'layout-0').theme.fonts.body, 'Arial');
  assert.equal(inherited.slides[0].backgroundSource, 'theme');
  assert.equal(Object.hasOwn(d.masters[0], 'theme'), false); assert.equal(Object.hasOwn(d.masters[0], 'background'), false);
});

test('new master declarations can repair a missing binding, with stable order and owned snapshots', () => {
  const d = createPresentation(); d.defineSlideLayout({ id: 'l', name: 'L', masterId: 'future' }); d.addSlide({ layoutId: 'l' });
  assert.throws(() => normalizePresentation(d), validation('missing-layout-master'));
  const input = { id: 'future', name: 'Future', theme: { colors: { accent1: '123456' } } };
  const m = d.defineSlideMaster(input); input.theme.colors.accent1 = 'FFFFFF';
  assert.equal(normalizePresentation(d).layouts.find(l => l.id === 'l').theme.colors.accent1, '123456');
  assert.equal(d.masters[0], m); assert.ok(Object.isFrozen(d.masters));
  assert.throws(() => d.defineSlideMaster({ id: 'future', name: 'Duplicate' }), /Duplicate/);
  d.defineSlideMaster({ id: 'second', name: 'Second' }); assert.deepEqual(d.masters.map(m => m.id), ['future', 'second']);
});

test('invalid master replacements leave the exact registry and document unchanged', () => {
  const d = multiMasterDocument(createPresentation), registry = d.masters, before = normalizePresentation(d);
  for (const input of [{ id: 'blue', name: 'Wrong ID' }, { name: '' }, { name: 'Bad', theme: { colors: { accent1: 'bad' } } }, { name: 'Bad', theme: { fonts: { body: '' } } }, { name: 'Bad', background: { type: 'solid', color: 'FFFFFF', opacity: -1 } }, { name: 'Bad', parentId: 'red' }, { name: 'Bad', elements: {} }]) {
    assert.throws(() => d.replaceSlideMaster('red', input)); assert.equal(d.masters, registry); assert.deepEqual(normalizePresentation(d), before);
  }
  assert.throws(() => d.replaceSlideMaster('absent', { name: 'Missing' }), /Unknown/);
});

test('layout full replacements retain shared identity, update all consumers and cancel background/master overrides', () => {
  const d = multiMasterDocument(createPresentation); d.addSlide({ id: 'extra', layoutId: 'layout-0', elements: [{ type: 'text', placeholderKey: 'same', content: 'Shared' }] });
  const old = d.layouts[0], before = normalizePresentation(d);
  const input = { name: 'Rebound', masterId: 'blue', placeholders: [placeholder()], background: { type: 'solid', color: '111111', opacity: 0 } };
  const replacement = d.replaceSlideLayout('layout-0', input); input.placeholders[0].box.x = 900;
  assert.equal(d.layouts[0], replacement); assert.equal(replacement.id, old.id); assert.equal(old.masterId, 'red'); assert.ok(Object.isFrozen(replacement.placeholders[0].box));
  let n = normalizePresentation(d);
  for (const i of [0, 4]) { assert.equal(n.slides[i].background.opacity, 0); assert.equal(n.slides[i].backgroundSource, 'layout'); assert.equal(n.slides[i].elements[0].box.x, 10); }
  for (const i of [1, 2, 3]) assert.deepEqual(n.slides[i], before.slides[i]);
  assert.equal(n.layouts.find(l => l.id === 'layout-0').theme.colors.accent1, '0044CC');
  d.replaceSlideLayout('layout-0', { name: 'Resume master', masterId: 'blue', placeholders: [placeholder()] });
  n = normalizePresentation(d); assert.equal(n.slides[0].backgroundSource, 'master');
  d.replaceSlideLayout('layout-0', { name: 'Default owner', placeholders: [placeholder()], background: { type: 'none' } });
  n = normalizePresentation(d); assert.equal(Object.hasOwn(n.layouts.find(l => l.id === 'layout-0'), 'masterId'), false); assert.equal(n.slides[0].background.type, 'none');
});

test('layout collection identity changes are atomic across sibling roots and nested groups', () => {
  const d = createPresentation();
  d.defineSlideLayout({ id: 'l', name: 'L', elements: [shape('a'), { id: 'g', type: 'group', box, coordinateSize: { width: 100, height: 50 }, children: [shape('b')] }] });
  const s = d.addSlide(); s.addElement(shape('external'));
  const old = d.layouts[0];
  assert.throws(() => d.replaceSlideLayout('l', { name: 'Collision', elements: [shape('a'), shape('external')] }), /Duplicate/);
  assert.equal(d.layouts[0], old); assert.throws(() => s.addElement(shape('b')), /Duplicate/);
  assert.throws(() => d.replaceSlideLayout('l', { name: 'Internal collision', elements: [shape('new'), shape('new')] }), /Duplicate/);
  s.addElement(shape('new')); // Failed transactions reserve no IDs.
  d.replaceSlideLayout('l', { name: 'Rearranged', elements: [shape('b'), shape('a')] });
  assert.deepEqual(d.layouts[0].elements.map(e => e.id), ['b', 'a']);
  s.addElement(shape('g')); // Removed group identity released only after commit.
  assert.throws(() => s.addElement(shape('a')), /Duplicate/);
  d.replaceSlideLayout('l', { name: 'Empty' }); s.addElement(shape('a')); s.addElement(shape('b'));
  assert.equal(validatePresentation(d).length, 0);
});

test('removing a shared connector target or placeholder aborts before IDs and dependencies change', () => {
  const d = createPresentation();
  d.defineSlideLayout({ id: 'l', name: 'L', elements: [shape('target')], placeholders: [placeholder()] });
  const s = d.addSlide({ layoutId: 'l', elements: [{ id: 'text', type: 'text', placeholderKey: 'same', content: 'Dependent' }, { id: 'link', type: 'connector', start: { elementId: 'target' }, end: { x: 400, y: 400 } }] });
  const old = d.layouts[0];
  assert.throws(() => d.replaceSlideLayout('l', { name: 'Break connector', elements: [shape('provisional')], placeholders: [placeholder()] }), validation('missing-connector-target'));
  assert.equal(d.layouts[0], old); s.addElement(shape('provisional'));
  assert.throws(() => d.replaceSlideLayout('l', { name: 'Break placeholder', elements: [shape('target')] }), validation('missing-placeholder'));
  assert.equal(d.layouts[0], old); assert.throws(() => s.addElement(shape('target')), /Duplicate/);
  d.replaceSlideLayout('l', { name: 'Valid', elements: [shape('target')], placeholders: [placeholder()] });
  assert.equal(validatePresentation(d).length, 0);
});

test('aborted layout replacements do not advance automatic identity allocation', () => {
  const d = createPresentation(); d.defineSlideLayout({ id: 'l', name: 'L', placeholders: [placeholder()] });
  const s = d.addSlide({ layoutId: 'l', elements: [{ id: 'body', type: 'text', content: 'Body', placeholderKey: 'same' }] });
  assert.throws(() => d.replaceSlideLayout('l', { name: 'Invalid', elements: [shape(undefined)] }), validation('missing-placeholder'));
  assert.equal(s.addElement(shape(undefined)).id, 'element-1');
  d.replaceSlideLayout('l', { name: 'Valid', placeholders: [placeholder()], elements: [shape(undefined)] });
  assert.equal(d.layouts[0].elements[0].id, 'element-2');
  assert.equal(s.addElement(shape(undefined)).id, 'element-3');
});

test('definition edits tolerate unrelated invalid drafts but reject affected invalid consumers and missing master references', () => {
  const d = createPresentation(); d.defineSlideLayout({ id: 'l', name: 'L' });
  const unrelated = d.addSlide({ layoutId: 'missing' });
  d.replaceSlideLayout('l', { name: 'Fine' }); d.defineSlideMaster({ id: 'm', name: 'M' });
  const old = d.layouts[0];
  assert.throws(() => d.replaceSlideLayout('l', { name: 'Missing master', masterId: 'missing' }), validation('missing-layout-master'));
  assert.equal(d.layouts[0], old);
  const affected = d.addSlide({ layoutId: 'l', elements: [{ type: 'text', content: 'x', placeholderKey: 'missing' }] });
  assert.throws(() => d.replaceSlideLayout('l', { name: 'Affected invalid' }), validation('missing-placeholder'));
  d.removeSlide(affected.id); d.removeSlide(unrelated.id);
  assert.equal(validatePresentation(d).length, 0);
});

test('layout replacement updates asset ownership only after a successful transaction', () => {
  const d = createPresentation();
  for (const id of ['logo', 'replacement']) d.registerAsset({ id, kind: 'image', source: { type: 'path', value: `./${id}.png` } });
  const image = (id, assetId) => ({ id, type: 'image', box, assetId });
  d.defineSlideLayout({ id: 'l', name: 'L', elements: [image('art', 'logo')] });
  const old = d.layouts[0];
  assert.throws(() => d.replaceSlideLayout('l', { name: 'Missing asset', elements: [image('art', 'missing')] }), validation('missing-image-asset'));
  assert.equal(d.layouts[0], old); assert.equal(d.getAssetReferences('logo').length, 1); assert.throws(() => d.removeAsset('logo'), /referenced/);
  d.replaceSlideLayout('l', { name: 'Updated image', elements: [image('art', 'replacement')] });
  assert.equal(d.getAssetReferences('logo').length, 0); assert.equal(d.getAssetReferences('replacement').length, 1);
  assert.deepEqual(d.removeUnusedAssets().map(a => a.id), ['logo']);
  assert.equal(validatePresentation(d).length, 0);
});

test('new definition edits reject executable, cyclic, sparse, over-depth and malformed inputs before committing', () => {
  const d = createPresentation({ masters: [{ id: 'm', name: 'M' }] }); d.defineSlideLayout({ id: 'l', name: 'L' });
  const registry = d.masters, layout = d.layouts[0]; let calls = 0;
  const accessor = Object.defineProperty({ name: 'Bad' }, 'theme', { enumerable: true, get() { calls++; d.defineSlideMaster({ id: 'reentrant', name: 'Bad' }); return {}; } });
  const cycle = { name: 'Cycle' }; cycle.theme = cycle;
  const symbol = { name: 'Symbol', [Symbol('hidden')]: 1 };
  const hidden = Object.defineProperty({ name: 'Hidden' }, 'theme', { value: {} });
  let deep = {}; for (let i = 0; i < 130; i++) deep = { next: deep };
  assert.throws(() => d.replaceSlideMaster('m', { name: 'Deep', theme: deep }), error => error instanceof RangeError && /depth 128/.test(error.message));
  assert.throws(() => d.replaceSlideMaster('m', cycle), /Cyclic definition/);
  for (const input of [accessor, symbol, hidden, { name: 'Date', theme: new Date() }, { name: 'Function', theme: () => {} }, { name: 'NaN', theme: { x: NaN } }, { name: 'Null background', background: null }]) assert.throws(() => d.replaceSlideMaster('m', input));
  for (const input of [{ name: 'Sparse', elements: new Array(2) }, { name: 'Wrong', elements: {} }, { name: 'Unknown', elements: [{ type: 'unknown' }] }, { name: 'Null', elements: [null] }, { name: 'Null background', background: null }, { name: 'Group', elements: [{ type: 'group', children: null }] }, { name: 'Wrong ID', id: 'other' }, { name: '' }, { name: 'Bad', parentId: 'l' }]) assert.throws(() => d.replaceSlideLayout('l', input));
  assert.throws(() => d.replaceSlideLayout('missing', { name: 'Missing' }), /Unknown/);
  assert.equal(calls, 0); assert.equal(d.masters, registry); assert.equal(d.layouts[0], layout);
  const bounded = createPresentation(); bounded.defineSlideLayout({ id: 'depth', name: 'Depth' });
  let customData = {}; for (let i = 0; i < 125; i++) customData = { next: customData };
  bounded.replaceSlideLayout('depth', { name: 'At depth 128', elements: [{ ...shape('at-limit'), customData }] });
  const atLimit = bounded.layouts[0];
  assert.throws(() => bounded.replaceSlideLayout('depth', { name: 'Depth 129', elements: [{ ...shape('over-limit'), customData: { next: customData } }] }), RangeError);
  assert.equal(bounded.layouts[0], atLimit);
});
