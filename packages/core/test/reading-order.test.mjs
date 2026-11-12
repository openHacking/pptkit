import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation, prepareLayoutRebind, validatePresentation } from '../dist/index.js';

const box = { x: 0, y: 0, width: 120, height: 50 };
const shape = id => ({ id, type: 'shape', shape: 'rect', box });
const group = (id, children) => ({ id, type: 'group', box, coordinateSize: box, children });
const elements = () => [shape('a'), group('g', [shape('b'), shape('c')])];
const orders = () => [{ elementIds: ['g', 'a'] }, { groupId: 'g', elementIds: ['c', 'b'] }];

test('authored order is captured, validated, and reconciled independently of drawing order', () => {
  const deck = createPresentation(), input = orders();
  assert.throws(() => deck.addSlide({ id: 's', elements: elements(), readingOrder: [{ elementIds: ['b', 'a'] }] }), /direct child/);
  assert.equal(deck.slides.length, 0);
  const slide = deck.addSlide({ id: 's', elements: elements(), readingOrder: input });
  input[1].elementIds.reverse();
  assert.deepEqual(slide.readingOrder, orders());
  assert.ok(Object.isFrozen(slide.readingOrder[1].elementIds));
  assert.throws(() => slide.setReadingOrder([{ groupId: 'g', elementIds: ['a', 'b'] }]), /direct child/);
  slide.moveElement('a', 1);
  slide.reorderGroupElements('g', ['c', 'b']);
  assert.deepEqual(slide.readingOrder, orders());
  slide.insertElement(0, shape('new'));
  assert.deepEqual(slide.readingOrder[0].elementIds, ['g', 'a', 'new']);
  slide.replaceElement('g', group('g', [{ ...shape('b'), hidden: true }, shape('d'), shape('c')]));
  assert.deepEqual(slide.readingOrder[1], { groupId: 'g', elementIds: ['c', 'd'] });
  slide.removeElement('a');
  assert.deepEqual(normalizePresentation(deck).slides[0].readingOrder, [{ elementIds: ['g', 'new'] }, { groupId: 'g', elementIds: ['c', 'd'] }]);
  slide.replaceElement('g', { ...group('g', [shape('c'), shape('d')]), accessibility: { decorative: true } });
  assert.deepEqual(slide.readingOrder, [{ elementIds: ['new'] }]);
  slide.removeElement('new');
  assert.equal(slide.readingOrder, undefined);
  assert.deepEqual(validatePresentation(deck), []);
});

test('whole-slide copies remap root, group and child identities without changing the source', () => {
  const deck = createPresentation();
  const slide = deck.addSlide({ id: 's', elements: elements(), readingOrder: orders() });
  const copy = deck.duplicateSlide(slide.id);
  for (const duplicate of [copy, createPresentation().copySlideFrom(deck, slide.id, {
    layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'reject', assets: { deduplicate: 'content', external: 'reject' },
  }).slide]) {
    const [a, g] = duplicate.elements, [b, c] = g.children;
    assert.deepEqual(duplicate.readingOrder, [{ elementIds: [g.id, a.id] }, { groupId: g.id, elementIds: [c.id, b.id] }]);
    assert.notEqual(g.id, 'g');
  }
  assert.deepEqual(slide.readingOrder, orders());
});

test('switching layout retains the independent order of slide-owned objects', () => {
  const deck = createPresentation();
  deck.defineSlideLayout({ id: 'old', name: 'Old', elements: [shape('old-art')] });
  deck.defineSlideLayout({ id: 'new', name: 'New', elements: [shape('new-art')] });
  const slide = deck.addSlide({ id: 's', layoutId: 'old', elements: elements(), readingOrder: orders() });
  const plan = prepareLayoutRebind(deck, [{ slideId: 's', layoutId: 'new' }], { localAppearance: 'adopt' });
  assert.equal(plan.status, 'ready');
  assert.equal(deck.applyLayoutRebind(plan).status, 'applied');
  assert.equal(slide.layoutId, 'new');
  assert.deepEqual(slide.readingOrder, orders());
  assert.deepEqual(normalizePresentation(deck).slides[0].elements.map(element => element.id), ['a', 'g']);
});
